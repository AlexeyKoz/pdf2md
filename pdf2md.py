"""pdf2md.py — converts PDF and Word (.docx) documents to Markdown (headings, tables, lists, images).
Core of the window (pdf2md_gui.py) and a command-line tool:

  py pdf2md.py                          every PDF/DOCX in input/  ->  output/
  py pdf2md.py FILE_OR_FOLDER ...       the given files / all PDF/DOCX in the given folders
Options:
  --out DIR          output folder (default: output/ next to this script)
  --no-images        do not save pictures          --dpi N        picture quality for PDFs (default 150)
  --keep-headers     keep page headers/footers     --keep-toc     keep the table of contents
  --no-numbering     do not restore Word heading numbers (1, 1.1 ...)
  --page-sep         mark page breaks (PDF) with ---      --version
"""
import os, re, sys, io, contextlib

__version__ = "1.0.0"
from collections import Counter
import pymupdf as fitz
import pymupdf4llm
if hasattr(pymupdf4llm, "use_layout"): pymupdf4llm.use_layout(False)   # 0.2+/1.x: classic engine (margins, no OCR)

HERE = os.path.dirname(os.path.abspath(__file__))
IN, OUT = os.path.join(HERE, "input"), os.path.join(HERE, "output")

def _norm(t):  # page numbers / dates differ between pages -> mask digits so repeated headers match
    return re.sub(r"\d+", "#", re.sub(r"\s+", " ", t)).strip().lower()

def detect_margins(doc, band=0.16, share=0.5):
    """finds repeated header/footer blocks (same text on >= half of the pages) and returns (top, bottom) in points"""
    n = len(doc)
    if n < 3: return 0, 0
    tops, bots, seen = Counter(), Counter(), []
    for p in doc:
        h = p.rect.height; blocks = [b for b in p.get_text("blocks") if b[4].strip()]
        t = {_norm(b[4]) for b in blocks if b[3] < h * band}; bt = {_norm(b[4]) for b in blocks if b[1] > h * (1 - band)}
        tops.update(t); bots.update(bt); seen.append((p, blocks))
    rep_t = {k for k, c in tops.items() if c >= n * share}; rep_b = {k for k, c in bots.items() if c >= n * share}
    top = bottom = 0
    for p, blocks in seen:
        h = p.rect.height
        for b in blocks:
            k = _norm(b[4])
            if k in rep_t and b[3] < h * band: top = max(top, b[3] + 2)
            if k in rep_b and b[1] > h * (1 - band): bottom = max(bottom, h - b[1] + 2)
    return top, bottom

def clean(md, strip_toc=True):
    md = md.replace("** **", " ")                                                   # "**Assembly** **Procedure**" -> one bold run
    md = re.sub(r"(?m)^\|Col\d+\|", "| |", md); md = re.sub(r"\|Col\d+(?=\|)", "| ", md)   # unnamed table columns
    if strip_toc: md = re.sub(r"(?m)^.*\.{6,}\s*\d+\s*$\n?", "", md)              # "3.1 Scope ........ 4" (table of contents)
    if strip_toc: md = re.sub(r"(?im)^#+\s*\**\s*(table of (contents|figures|tables)|contents|תוכן העניינים|תוכן עניינים|содержание|оглавление)\s*\**\s*$\n?", "", md)
    md = re.sub(r"(?im)^\s*(page|עמוד|стр\.?)\s*\d+\s*(of|/|из)\s*\d+\s*$", "", md)   # "Page 3 of 15"
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip() + "\n"

SYMBOL = dict(zip("abgdezhqiklmnxprstufcywDGQLXPSFYW", "αβγδεζηθικλμνξπρστυφχψωΔΓΘΛΞΠΣΦΨΩ")); SYMBOL["m"] = "µ"

def _prep_docx(path):
    """in-memory fixes of the Word XML before conversion:
    - text typed in the Symbol font ('m' shown as µ) -> real Unicode letters
    - simple fields (caption numbers 'Figure 3', 'Table 1') -> keep their shown text
    returns (BytesIO of fixed docx, headings_are_numbered)"""
    import zipfile
    zin = zipfile.ZipFile(path); buf = io.BytesIO(); zout = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith("word/") and item.filename.endswith(".xml") and item.filename.split("/")[-1].startswith(("document", "footnotes", "endnotes")):
            x = data.decode("utf-8")
            x = re.sub(r"<w:fldSimple\b[^>]*/>", "", x); x = re.sub(r"<w:fldSimple\b[^>]*>(.*?)</w:fldSimple>", r"\1", x, flags=re.S)
            def fix_run(m):
                r = m.group(0)
                if 'w:ascii="Symbol"' not in r: return r
                return re.sub(r"(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)", lambda t: t.group(1) + "".join(SYMBOL.get(c, c) for c in t.group(2)) + t.group(3), r)
            x = re.sub(r"<w:r\b(?:(?!<w:r\b).)*?</w:r>", fix_run, x, flags=re.S)
            x = re.sub(r'<w:sym w:font="Symbol" w:char="F0([0-9A-Fa-f]{2})"/>', lambda m: "<w:t>" + SYMBOL.get(chr(int(m.group(1), 16)), chr(int(m.group(1), 16))) + "</w:t>", x)
            data = x.encode("utf-8")
        zout.writestr(item.filename, data)
    numbered = False
    try:
        st = zin.read("word/styles.xml").decode("utf-8")
        m = re.search(r'<w:style [^>]*w:styleId="Heading1".*?</w:style>', st, re.S)
        numbered = bool(m and "<w:numPr>" in m.group(0))
    except KeyError: pass
    zout.close(); buf.seek(0)
    return buf, numbered

def _number_headings(md):
    """Word numbers headings automatically (1, 1.1, 1.1.1); the numbers are not in the text, so restore them"""
    cnt = [0] * 7; out = []
    for line in md.split("\n"):
        m = re.match(r"^(#{1,4}) (.*)", line)
        if m and not re.match(r"^\**\s*\d", m.group(2)):
            lvl = len(m.group(1)); cnt[lvl] += 1; cnt[lvl + 1:] = [0] * (6 - lvl)
            line = f"{m.group(1)} {'.'.join(str(c) for c in cnt[1:lvl + 1])} {m.group(2)}"
        out.append(line)
    return "\n".join(out)

STYLE_MAP = """
p[style-name^='Heading 1'] => h1:fresh
p[style-name^='Heading 2'] => h2:fresh
p[style-name^='Heading 3'] => h3:fresh
p[style-name^='Heading 4'] => h4:fresh
p[style-name='caption'] => p:fresh
"""
TOC_MAP = """
p[style-name^='toc '] => !
p[style-name='table of figures'] => !
"""

def _emf_to_png(img_dir, fn, log):
    """EMF/WMF are not shown by browsers/markdown viewers; Pillow can render them on Windows"""
    try:
        from PIL import Image
        src = os.path.join(img_dir, fn); dst = os.path.splitext(src)[0] + ".png"
        with Image.open(src) as im: im.load(dpi=150) if hasattr(im, "load") else None; im.save(dst)
        os.remove(src); return os.path.basename(dst)
    except Exception:
        return fn

def convert_docx(path, out_dir=OUT, images=True, strip_toc=True, number_headings=True, log=print, name=None, **_):
    """Word .docx -> Markdown via mammoth (docx -> HTML) + markdownify (HTML -> md).
    Headers/footers and review comments are not part of the body, so they are skipped automatically."""
    import mammoth, markdownify
    os.makedirs(out_dir, exist_ok=True)
    name = name or os.path.splitext(os.path.basename(path))[0]; img_dir = os.path.join(out_dir, name + "_images"); cnt = [0]; bad = [0]
    def save_img(image):
        if not images: return {"src": ""}
        os.makedirs(img_dir, exist_ok=True); cnt[0] += 1
        ext = (image.content_type or "image/png").split("/")[-1].replace("jpeg", "jpg").replace("x-emf", "emf").replace("x-wmf", "wmf")
        fn = f"{name}-{cnt[0]}.{ext}"
        with image.open() as src, open(os.path.join(img_dir, fn), "wb") as dst: dst.write(src.read())
        if ext in ("emf", "wmf"):
            fn = _emf_to_png(img_dir, fn, log)
            if fn.endswith(("emf", "wmf")): bad[0] += 1
        return {"src": f"{name}_images/{fn}", "alt": ""}
    fixed, numbered = _prep_docx(path)
    res = mammoth.convert_to_html(fixed, convert_image=mammoth.images.img_element(save_img),
                                  style_map=STYLE_MAP + (TOC_MAP if strip_toc else ""))
    md = markdownify.markdownify(res.value, heading_style="ATX", bullets="-", strip=["a"])
    md = re.sub(r"!\[[^\]]*\]\(\)", "", md)                                           # images switched off
    md = re.sub(r"(?m)^(\|(?:\s*\|)+)\n(\|(?:\s*-+\s*\|)+)\n(\|.*\|)$", r"\3\n\2", md)  # Word tables: empty header row
    if numbered and number_headings: md = _number_headings(md); log("   heading numbers restored (1, 1.1, …)")
    md = clean(md, strip_toc)
    if strip_toc: md = re.sub(r"(?im)^\**\s*(table of (contents|figures|tables)|contents)\s*\**\s*$\n?", "", md)
    md = re.sub(r"\n{3,}", "\n\n", md)
    if cnt[0]: log(f"   pictures saved: {cnt[0]}" + (f"  ({bad[0]} EMF/WMF could not be converted to PNG on this computer)" if bad[0] else ""))
    out = os.path.join(out_dir, name + ".md")
    with open(out, "w", encoding="utf-8") as f: f.write(md)
    return out

SUPPORTED = (".pdf", ".docx")

def convert_any(path, **kw):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx": return convert_docx(path, **kw)
    if ext == ".pdf": return convert(path, **kw)
    if ext == ".doc": raise RuntimeError("old .doc format - open it in Word and save as .docx")
    raise RuntimeError(f"{ext} files are not supported (PDF and DOCX only)")

def convert(pdf_path, out_dir=OUT, images=True, strip_hf=True, strip_toc=True, page_sep=False, dpi=150, log=print, name=None, **_):
    """converts one PDF, returns path of the .md file"""
    os.makedirs(out_dir, exist_ok=True)
    name = name or os.path.splitext(os.path.basename(pdf_path))[0]
    doc = fitz.open(pdf_path)
    if doc.needs_pass: raise RuntimeError("the PDF is password-protected")
    top, bottom = detect_margins(doc) if strip_hf else (0, 0)
    if strip_hf and (top or bottom): log(f"   page header {top:.0f} pt / footer {bottom:.0f} pt removed")
    kw = dict(margins=(0, top, 0, bottom), page_chunks=True, show_progress=False)
    if images:
        img_dir = os.path.join(out_dir, name + "_images"); os.makedirs(img_dir, exist_ok=True)
        kw.update(write_images=True, image_path=img_dir, image_format="png", dpi=dpi)
    with contextlib.redirect_stdout(io.StringIO()): chunks = pymupdf4llm.to_markdown(doc, **kw)   # silence library hints
    parts = []
    for i, ch in enumerate(chunks, 1):
        t = ch["text"].strip()
        if images:  # make image links relative to the .md file and portable (forward slashes)
            t = t.replace(img_dir.replace("\\", "/"), name + "_images").replace(img_dir, name + "_images").replace("\\", "/") if "_images" in t else t
        parts.append(f"<!-- page {i} -->\n\n{t}" if page_sep else t)
    md = clean(("\n\n---\n\n" if page_sep else "\n\n").join(p for p in parts if p), strip_toc)
    if not md.strip():
        log("   WARNING: no text found - probably a scan (pages are pictures without a text layer)")
    if images and not os.listdir(img_dir): os.rmdir(img_dir)
    out = os.path.join(out_dir, name + ".md")
    with open(out, "w", encoding="utf-8") as f: f.write(md)
    return out

def collect(paths):
    """files and folders -> list of PDF/DOCX files (Word lock files ~$ skipped)"""
    out = []
    for p in paths:
        if os.path.isdir(p): out += [os.path.join(p, f) for f in sorted(os.listdir(p)) if f.lower().endswith(SUPPORTED) and not f.startswith("~$")]
        elif p.lower().endswith(SUPPORTED + (".doc",)): out.append(p)
        else: print(f"skipped (not PDF/DOCX): {p}")
    return out

def output_names(files):
    """file -> output name; "Manual.pdf" + "Manual.docx" in one run become Manual_pdf.md / Manual_docx.md"""
    stems = Counter(os.path.splitext(os.path.basename(f))[0].lower() for f in files)
    out = {}
    for f in files:
        stem, ext = os.path.splitext(os.path.basename(f))
        out[f] = f"{stem}_{ext[1:].lower()}" if stems[stem.lower()] > 1 else stem
    return out

if __name__ == "__main__":   # files/folders on the command line (or dragged onto run.bat), else everything in input/
    A = sys.argv[1:]
    if "--version" in A: sys.exit(f"pdf2md {__version__}")
    if "-h" in A or "--help" in A: sys.exit(__doc__)
    def opt(k, d=None): return A[A.index(k) + 1] if k in A else d
    vals = {opt("--out"), opt("--dpi")}
    paths = [a for a in A if not a.startswith("--") and a not in vals]
    os.makedirs(IN, exist_ok=True)
    files = collect(paths or [IN])
    if not files: print("No PDF/DOCX files found" + ("" if paths else " in the input folder") + "."); sys.exit(0)
    opts = dict(out_dir=opt("--out", OUT), images="--no-images" not in A, strip_hf="--keep-headers" not in A,
                strip_toc="--keep-toc" not in A, number_headings="--no-numbering" not in A, page_sep="--page-sep" in A,
                dpi=int(opt("--dpi", 150)))
    ok = 0; names = output_names(files)
    for f in files:
        print(f"-> {os.path.basename(f)}")
        try: out = convert_any(f, name=names[f], **opts); print(f"OK  {os.path.basename(out)}"); ok += 1
        except Exception as e: print(f"ERROR: {e}")
    print(f"\nDone: {ok} of {len(files)} converted - see {opts['out_dir']}")
    sys.exit(0 if ok == len(files) else 1)
