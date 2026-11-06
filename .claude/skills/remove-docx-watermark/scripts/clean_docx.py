#!/usr/bin/env python3
"""Find and remove AI-generation fingerprints from a Word (.docx) file.

Works offline with the Python standard library plus lxml. Nothing is uploaded.

Usage:
    python clean_docx.py inspect  report.docx
    python clean_docx.py clean    report.docx -o report.clean.docx --author "Your Name"

What it handles (see SKILL.md for the reasoning behind each item):
  1. Core/app/custom properties  (creator "Claude", "python-docx", "Un-named",
     millisecond timestamps, empty app.xml, AI-flavoured custom properties)
  2. Review authors              (comments / tracked changes by "Claude", "AI", ...)
  3. Visible watermarks          (VML/WordArt or DrawingML watermark shapes in headers)
  4. Invisible Unicode           (zero-width, BOM, bidi controls, tag characters, odd spaces)
  5. Missing Word save history   (adds w:rsid revision ids that Word always writes)
  6. Embedded image metadata     (PNG text/XMP/C2PA chunks, JPEG EXIF/XMP/C2PA segments)
  7. Optional em-dash softening  (--dashes en|comma)

It does NOT change wording. Statistical "style" watermarks live in the prose
itself and can only be reduced by rewriting the text.
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import random
import re
import struct
import sys
import zipfile

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    sys.exit("lxml is required: pip install lxml")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {
    "w": W,
    "v": "urn:schemas-microsoft-com:vml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

AI_TERMS = re.compile(
    r"claude|anthropic|chat\s*gpt|openai|\bgpt[-\s]?\d|gemini|copilot|bard\b|llama|"
    r"mistral|\bllm\b|\bA\.?I\.?\b|artificial intelligence|machine[- ]generated|"
    r"generated (by|with|using)|python-docx|docx-?js|un-named|unnamed",
    re.IGNORECASE,
)
# Authors that tools write by default; replaced on comments / tracked changes.
AI_AUTHORS = re.compile(r"^(claude|anthropic|ai|assistant|chatgpt|gpt|copilot|gemini|un-named|author)$", re.I)

# Characters with no visible purpose in normal prose.
INVISIBLE = dict.fromkeys(
    [0x200B, 0x200C, 0x200D, 0x2060, 0x2061, 0x2062, 0x2063, 0x2064, 0xFEFF, 0x00AD,
     0x180E, 0x034F, 0x061C, 0x200E, 0x200F, 0x202A, 0x202B, 0x202C, 0x202D, 0x202E,
     0x2066, 0x2067, 0x2068, 0x2069]
    + list(range(0xE0000, 0xE0080))   # Unicode tag characters
    + list(range(0xFE00, 0xFE10))     # variation selectors
)
ODD_SPACES = {0x2000, 0x2001, 0x2002, 0x2003, 0x2004, 0x2005, 0x2006, 0x2007,
              0x2008, 0x2009, 0x200A, 0x202F, 0x205F, 0x3000}

WATERMARK_HINT = re.compile(r"PowerPlusWaterMarkObject|PowerPlusPictureWaterMark|WordPictureWatermark|watermark", re.I)

TEXT_PARTS = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments\w*)\.xml$")
AUTHOR_PARTS = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments\w*|people)\.xml$")


# --------------------------------------------------------------------------- helpers
def parse(data: bytes):
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)
    return etree.fromstring(data, parser)


def dump(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def wq(tag: str) -> str:
    return f"{{{W}}}{tag}"


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0).strftime("%Y-%m-%dT%H:%M:%SZ")


def strip_ms(value: str) -> str:
    return re.sub(r"\.\d+(?=Z|[+-]\d\d:?\d\d$)", "", value.strip())


def clean_string(s: str, dashes: str, keep_nbsp: bool, counts: dict) -> str:
    out = []
    for ch in s:
        cp = ord(ch)
        if cp in INVISIBLE:
            counts["invisible"] = counts.get("invisible", 0) + 1
            continue
        if cp in ODD_SPACES or (cp == 0x00A0 and not keep_nbsp):
            counts["spaces"] = counts.get("spaces", 0) + 1
            out.append(" ")
            continue
        out.append(ch)
    s = "".join(out)
    if dashes != "keep" and "—" in s:
        counts["dashes"] = counts.get("dashes", 0) + s.count("—")
        if dashes == "en":
            s = re.sub(r"\s*—\s*", " – ", s)
        else:  # comma
            s = re.sub(r"\s*—\s*", ", ", s)
    return s


# --------------------------------------------------------------------------- images
PNG_DROP = {b"tEXt", b"iTXt", b"zTXt", b"eXIf", b"caBX", b"tIME"}


def strip_png(data: bytes) -> tuple[bytes, int]:
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        return data, 0
    out, pos, dropped = [data[:8]], 8, 0
    while pos + 8 <= len(data):
        length, ctype = struct.unpack(">I4s", data[pos:pos + 8])
        end = pos + 12 + length
        if ctype in PNG_DROP:
            dropped += 1
        else:
            out.append(data[pos:end])
        pos = end
        if ctype == b"IEND":
            break
    return b"".join(out), dropped


def strip_jpeg(data: bytes) -> tuple[bytes, int]:
    """Drop APP1 (EXIF/XMP), APP11 (C2PA/JUMBF), APP12, APP13 (IPTC) and COM segments."""
    if not data.startswith(b"\xff\xd8"):
        return data, 0
    out, pos, dropped = [b"\xff\xd8"], 2, 0
    while pos + 4 <= len(data):
        if data[pos] != 0xFF:
            return data, 0  # unexpected layout: leave untouched
        marker = data[pos + 1]
        if marker == 0xDA:  # start of scan: copy the rest verbatim
            out.append(data[pos:])
            break
        if 0xD0 <= marker <= 0xD7 or marker == 0x01:
            out.append(data[pos:pos + 2]); pos += 2; continue
        seglen = struct.unpack(">H", data[pos + 2:pos + 4])[0]
        end = pos + 2 + seglen
        if marker in (0xE1, 0xEB, 0xEC, 0xED, 0xFE):
            dropped += 1
        else:
            out.append(data[pos:end])
        pos = end
    return b"".join(out), dropped


# --------------------------------------------------------------------------- inspect
def inspect(path: str) -> list[str]:
    findings: list[str] = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if any(n.endswith("/") for n in names):
            findings.append("ZIP has folder entries (Word never writes these)")
        if "docProps/core.xml" in names:
            core = parse(z.read("docProps/core.xml"))
            for el in core:
                tag = etree.QName(el).localname
                text = (el.text or "").strip()
                if text and AI_TERMS.search(text):
                    findings.append(f"core.xml <{tag}> = {text!r}")
                if tag in ("created", "modified") and re.search(r"\.\d+Z", text):
                    findings.append(f"core.xml <{tag}> has milliseconds ({text}); Word writes whole seconds")
            c = core.find("dcterms:created", NS); m = core.find("dcterms:modified", NS)
            if c is not None and m is not None and c.text == m.text:
                findings.append("created == modified to the millisecond (file was never re-saved by a person)")
        if "docProps/app.xml" in names:
            app = z.read("docProps/app.xml").decode("utf8", "replace")
            if "<Application>" not in app:
                findings.append("app.xml has no <Application> (Word always records itself)")
            elif not re.search(r"<Application>Microsoft", app):
                findings.append("app.xml Application = " + re.search(r"<Application>(.*?)</", app).group(1))
        else:
            findings.append("docProps/app.xml missing")
        if "docProps/custom.xml" in names:
            cust = z.read("docProps/custom.xml").decode("utf8", "replace")
            if "<property" not in cust:
                findings.append("empty docProps/custom.xml part (docx-js signature)")
            elif AI_TERMS.search(cust):
                findings.append("custom.xml contains AI-related properties")
        rsid = False
        inv = spaces = dashes = 0
        for n in names:
            if not n.endswith(".xml"):
                continue
            raw = z.read(n)
            txt = raw.decode("utf8", "replace")
            if AUTHOR_PARTS.match(n):
                for a in set(re.findall(r'w:author="([^"]*)"', txt)) | set(re.findall(r'w15:author="([^"]*)"', txt)):
                    if AI_AUTHORS.match(a) or AI_TERMS.search(a):
                        findings.append(f"{n}: review author {a!r}")
            if re.match(r"word/header\d*\.xml", n) and WATERMARK_HINT.search(txt):
                findings.append(f"{n}: visible watermark shape")
            if TEXT_PARTS.match(n):
                if "w:rsidR" in txt:
                    rsid = True
                for ch in txt:
                    cp = ord(ch)
                    if cp in INVISIBLE:
                        inv += 1
                    elif cp in ODD_SPACES or cp == 0xA0:
                        spaces += 1
                    elif cp == 0x2014:
                        dashes += 1
        if not rsid:
            findings.append("no w:rsid revision ids (every Word-saved file has them; generators don't)")
        if inv:
            findings.append(f"{inv} invisible Unicode characters in text")
        if spaces:
            findings.append(f"{spaces} non-standard space characters (NBSP / thin / narrow)")
        if dashes:
            findings.append(f"{dashes} em dashes (a common AI-style tell; optional --dashes en)")
        for n in names:
            if n.startswith("word/media/"):
                data = z.read(n)
                if n.lower().endswith(".png") and strip_png(data)[1]:
                    findings.append(f"{n}: PNG metadata chunks (text/XMP/C2PA)")
                if n.lower().endswith((".jpg", ".jpeg")) and strip_jpeg(data)[1]:
                    findings.append(f"{n}: JPEG metadata segments (EXIF/XMP/C2PA)")
                if b"c2pa" in data[:65536].lower() or b"jumb" in data[:65536]:
                    findings.append(f"{n}: C2PA / Content Credentials manifest")
    return findings


# --------------------------------------------------------------------------- clean
def initials_of(name: str) -> str:
    return "".join(p[0] for p in name.split() if p[:1].isalpha()).upper()[:4] or "A"


def clean(src: str, dst: str, author: str, dashes: str, keep_nbsp: bool,
          strip_comments: bool, add_rsids: bool) -> list[str]:
    log: list[str] = []
    counts: dict = {}
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist() if not n.endswith("/")}
    zin.close()

    # ---- 1. core properties
    if "docProps/core.xml" in files:
        core = parse(files["docProps/core.xml"])
        for el in list(core):
            tag = etree.QName(el).localname
            text = (el.text or "").strip()
            if tag in ("creator", "lastModifiedBy"):
                if el.text != author:
                    log.append(f"core: {tag} {text!r} -> {author!r}")
                el.text = author
            elif tag in ("description", "subject", "keywords", "category", "contentStatus", "identifier") \
                    and text and AI_TERMS.search(text):
                core.remove(el); log.append(f"core: removed {tag} {text!r}")
            elif tag == "title" and text and AI_TERMS.search(text):
                el.text = re.sub(AI_TERMS, "", text).strip(" -:|"); log.append("core: AI words removed from title")
            elif tag == "created" and text:
                el.text = strip_ms(text)
            elif tag == "modified":
                el.text = now_iso()
            elif tag == "revision":
                try:
                    if int(text or 1) < 2:
                        el.text = str(random.randint(2, 6))
                except ValueError:
                    el.text = "2"
        if core.find("dc:creator", NS) is None:
            e = etree.SubElement(core, f"{{{NS['dc']}}}creator"); e.text = author
        if core.find("cp:lastModifiedBy", NS) is None:
            e = etree.SubElement(core, f"{{{NS['cp']}}}lastModifiedBy"); e.text = author
        files["docProps/core.xml"] = dump(core)

    # ---- 4. text cleaning (+ count words for app.xml)
    words = chars = 0
    for n in list(files):
        if not TEXT_PARTS.match(n):
            continue
        root = parse(files[n])
        for t in root.iter(wq("t"), wq("delText")):
            if t.text:
                t.text = clean_string(t.text, dashes, keep_nbsp, counts)
                if n == "word/document.xml" and t.tag == wq("t"):
                    words += len(t.text.split()); chars += len(t.text.replace(" ", ""))
        files[n] = dump(root)

    # ---- 2. review authors / comments
    for n in list(files):
        if not AUTHOR_PARTS.match(n):
            continue
        root = parse(files[n])
        changed = 0
        for el in root.iter():
            for attr in list(el.attrib):
                local = etree.QName(attr).localname
                if local == "author" and (AI_AUTHORS.match(el.get(attr)) or AI_TERMS.search(el.get(attr))):
                    el.set(attr, author or "Author"); changed += 1
                    ini = el.get(wq("initials"))
                    if ini is not None:
                        el.set(wq("initials"), initials_of(author or "Author"))
                elif local == "userId" and AI_TERMS.search(el.get(attr) or ""):
                    el.set(attr, author or "Author")
        if changed:
            log.append(f"{n}: {changed} review author(s) -> {author or 'Author'!r}")
        files[n] = dump(root)

    if strip_comments and "word/document.xml" in files:
        root = parse(files["word/document.xml"])
        removed = 0
        for tag in ("commentRangeStart", "commentRangeEnd"):
            for el in list(root.iter(wq(tag))):
                el.getparent().remove(el); removed += 1
        for ref in list(root.iter(wq("commentReference"))):
            run = ref.getparent()
            if run.tag == wq("r"):
                run.getparent().remove(run)
            else:
                run.remove(ref)
        files["word/document.xml"] = dump(root)
        for n in [k for k in files if re.match(r"word/comments\w*\.xml$", k)]:
            r = parse(files[n])
            for child in list(r):
                r.remove(child)
            files[n] = dump(r)
        if removed:
            log.append(f"comments: removed {removed // 2} comment(s)")

    # ---- 3. visible watermarks in headers
    for n in [k for k in files if re.match(r"word/header\d*\.xml$", k)]:
        root = parse(files[n])
        hits = 0
        for run in list(root.iter(wq("r"))):
            if run.getparent() is None:
                continue
            blob = etree.tostring(run).decode("utf8", "replace")
            is_vml_wm = (run.find(".//v:shape", NS) is not None and
                         (WATERMARK_HINT.search(blob) or run.find(".//v:textpath", NS) is not None))
            is_dml_wm = any(WATERMARK_HINT.search(d.get("name", "") + d.get("descr", ""))
                            for d in run.iterfind(".//wp:docPr", NS))
            if is_vml_wm or is_dml_wm:
                run.getparent().remove(run); hits += 1
        if hits:
            log.append(f"{n}: removed {hits} watermark shape(s)")
            files[n] = dump(root)

    # ---- 5. Word save history (rsids)
    if add_rsids and "word/settings.xml" in files:
        doc_has = b"w:rsidR" in files.get("word/document.xml", b"")
        if not doc_has:
            ids = sorted({f"{random.randint(0x00100000, 0x00FFFFFF):08X}" for _ in range(4)})
            root_id = ids[0]
            settings = parse(files["word/settings.xml"])
            if settings.find("w:rsids", NS) is None:
                rs = etree.Element(wq("rsids"))
                etree.SubElement(rs, wq("rsidRoot")).set(wq("val"), root_id)
                for i in ids:
                    etree.SubElement(rs, wq("rsid")).set(wq("val"), i)
                after = ["mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
                         "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade",
                         "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary",
                         "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"]
                anchor = next((c for c in settings if etree.QName(c).localname in after), None)
                if anchor is not None:
                    anchor.addprevious(rs)
                else:
                    settings.append(rs)
                files["word/settings.xml"] = dump(settings)
            for n in list(files):
                if not TEXT_PARTS.match(n) or n.startswith("word/comments"):
                    continue
                r = parse(files[n])
                for p in r.iter(wq("p")):
                    if p.get(wq("rsidR")) is None:
                        rid = random.choice(ids)
                        p.set(wq("rsidR"), rid); p.set(wq("rsidRDefault"), rid)
                files[n] = dump(r)
            log.append(f"settings: added Word revision history (rsidRoot {root_id})")

    # ---- 1b. app.xml (Word identity)
    app = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        f"<Template>Normal.dotm</Template><TotalTime>{random.randint(35, 240)}</TotalTime>"
        f"<Words>{words}</Words><Characters>{chars}</Characters>"
        "<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity>"
        "<ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate>"
        f"<CharactersWithSpaces>{chars + words}</CharactersWithSpaces><SharedDoc>false</SharedDoc>"
        "<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>"
    ).encode()
    files["docProps/app.xml"] = app
    log.append("app.xml: Application set to Microsoft Office Word 16")

    # ---- 1c. custom.xml
    if "docProps/custom.xml" in files:
        cust = parse(files["docProps/custom.xml"])
        for prop in list(cust):
            blob = etree.tostring(prop).decode("utf8", "replace")
            if AI_TERMS.search(blob):
                cust.remove(prop); log.append(f"custom: removed property {prop.get('name')!r}")
        if len(cust) == 0:
            del files["docProps/custom.xml"]
            log.append("custom: removed empty custom.xml part")
        else:
            files["docProps/custom.xml"] = dump(cust)
    ensure_package_parts(files)

    # ---- 6. images
    for n in list(files):
        if not n.startswith("word/media/"):
            continue
        low = n.lower()
        if low.endswith(".png"):
            files[n], d = strip_png(files[n])
        elif low.endswith((".jpg", ".jpeg")):
            files[n], d = strip_jpeg(files[n])
        else:
            d = 0
        if d:
            log.append(f"{n}: removed {d} metadata block(s)")

    for k, v in counts.items():
        label = {"invisible": "invisible characters removed", "spaces": "odd spaces normalised",
                 "dashes": f"em dashes replaced ({dashes})"}[k]
        log.append(f"text: {v} {label}")

    # ---- write package the way Word does: content types first, no folder entries
    order = ["[Content_Types].xml", "_rels/.rels"] + sorted(k for k in files if k not in ("[Content_Types].xml", "_rels/.rels"))
    stamp = dt.datetime.now().timetuple()[:6]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in order:
            if n in files:
                info = zipfile.ZipInfo(n, date_time=stamp)
                info.compress_type = zipfile.ZIP_DEFLATED
                zout.writestr(info, files[n])
    return log


def ensure_package_parts(files: dict) -> None:
    """Keep [Content_Types].xml and _rels/.rels consistent with app.xml / custom.xml."""
    ct = parse(files["[Content_Types].xml"])
    rels = parse(files["_rels/.rels"])
    want = {
        "/docProps/app.xml": ("application/vnd.openxmlformats-officedocument.extended-properties+xml",
                              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"),
    }
    for part, (ctype, rtype) in want.items():
        if not any(o.get("PartName") == part for o in ct.iter(f"{{{CT_NS}}}Override")):
            e = etree.SubElement(ct, f"{{{CT_NS}}}Override"); e.set("PartName", part); e.set("ContentType", ctype)
        if not any(r.get("Target", "").lstrip("/") == part.lstrip("/") for r in rels):
            ids = {r.get("Id") for r in rels}
            i = 1
            while f"rId{i}" in ids:
                i += 1
            e = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
            e.set("Id", f"rId{i}"); e.set("Type", rtype); e.set("Target", part.lstrip("/"))
    if "docProps/custom.xml" not in files:
        for o in list(ct.iter(f"{{{CT_NS}}}Override")):
            if o.get("PartName") == "/docProps/custom.xml":
                ct.remove(o)
        for r in list(rels):
            if r.get("Target", "").lstrip("/") == "docProps/custom.xml":
                rels.remove(r)
    files["[Content_Types].xml"] = dump(ct)
    files["_rels/.rels"] = dump(rels)


# --------------------------------------------------------------------------- cli
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    i = sub.add_parser("inspect", help="list fingerprints without changing anything")
    i.add_argument("docx")
    c = sub.add_parser("clean", help="write a cleaned copy")
    c.add_argument("docx")
    c.add_argument("-o", "--output", help="output path (default: <name>.clean.docx)")
    c.add_argument("--author", default="", help="name to record as author / reviewer (default: blank)")
    c.add_argument("--dashes", choices=["keep", "en", "comma"], default="keep",
                   help="em dash handling: keep, spaced en dash, or comma (default keep)")
    c.add_argument("--keep-nbsp", action="store_true", help="leave non-breaking spaces alone")
    c.add_argument("--strip-comments", action="store_true", help="delete all comments")
    c.add_argument("--no-rsids", action="store_true", help="do not add Word revision ids")
    a = ap.parse_args()

    if a.cmd == "inspect":
        f = inspect(a.docx)
        print(f"Fingerprints in {a.docx}: {len(f)}")
        for line in f:
            print("  -", line)
        return 1 if f else 0

    out = a.output or re.sub(r"\.docx$", "", a.docx, flags=re.I) + ".clean.docx"
    log = clean(a.docx, out, a.author, a.dashes, a.keep_nbsp, a.strip_comments, not a.no_rsids)
    print(f"Wrote {out}")
    for line in log:
        print("  -", line)
    left = [x for x in inspect(out) if "em dashes" not in x or a.dashes != "keep"]
    left = [x for x in left if not (a.keep_nbsp and "non-standard space" in x)]
    print(f"Re-inspection: {len(left)} fingerprint(s) left")
    for line in left:
        print("  -", line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
