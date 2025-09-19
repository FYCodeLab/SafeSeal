# app.py — SafeSeal v4 · Streamlit + LibreOffice + Watermark
import io, os, shutil, subprocess, tempfile, time, pathlib
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import fitz  # PyMuPDF

# ---------------------------
# Page setup
# ---------------------------
st.set_page_config(page_title="SafeSeal v4 · Streamlit + LibreOffice + Watermark",
                   layout="centered")
st.title("SafeSeal v4 · Streamlit + LibreOffice + Watermark")

# ---------------------------
# CSS — dark, fixed-height (4 lines) status box under “Status”
# ---------------------------
st.markdown("""
<style>
.status-box {
  background: #2b2b2b; color: #e6e6e6;
  border: 1px solid #3a3a3a; border-radius: 6px;
  padding: 8px 10px;
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
  font-size: 12px; line-height: 1.25;
  white-space: pre-wrap;
  height: calc(1.25em * 4 + 16px);  /* exactly 4 lines + padding */
  overflow-y: auto; overflow-x: hidden;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# LibreOffice detection
# ---------------------------
def _resolve_libreoffice_bin():
    for cand in ["soffice", "libreoffice", "/usr/bin/soffice", "/usr/bin/libreoffice"]:
        path = shutil.which(cand) if os.path.basename(cand) == cand else (cand if os.path.exists(cand) else None)
        if path:
            try:
                subprocess.run([path, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
                return path
            except Exception:
                continue
    return None

LO_BIN = _resolve_libreoffice_bin()

# ---------------------------
# Logging helpers (fixed-height box)
# ---------------------------
_status_box = st.empty()

def _render_status():
    buf = st.session_state.get("_logbuf", "")
    _status_box.markdown(f"<div class='status-box'>{buf}</div>", unsafe_allow_html=True)

def clear_log():
    st.session_state["_logbuf"] = ""
    _render_status()

def log_line(msg: str):
    buf = st.session_state.get("_logbuf", "")
    st.session_state["_logbuf"] = buf + msg.rstrip() + "\n"
    _render_status()

# ---------------------------
# Watermark helpers
# ---------------------------
def _load_font(px: int):
    try:
        return ImageFont.truetype("DejaVuSans.ttf", px)
    except Exception:
        return ImageFont.load_default()

def _draw_tiled_watermark(img_rgba, text, dpi=120, angle=45, opacity=60):
    w, h = img_rgba.size
    font_px = max(6, int(round(8 * dpi / 72.0)))   # ~8pt scaled by DPI
    font = _load_font(font_px)
    spacing_px = int(dpi)  # 1 inch
    layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    fill = (180, 180, 180, max(0, min(255, opacity)))
    for y in range(-h, h*2, spacing_px):
        for x in range(-w, w*2, spacing_px):
            draw.text((x, y), text, font=font, fill=fill)
    layer = layer.rotate(angle, expand=True, resample=Image.BICUBIC)
    lw, lh = layer.size
    left, top = (lw - w) // 2, (lh - h) // 2
    layer = layer.crop((left, top, left + w, top + h))
    return Image.alpha_composite(img_rgba, layer)

def pdf_to_imageonly_pdf_with_watermark(pdf_bytes, wm_text, dpi, quality, progress_cb=None, log_cb=None):
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()
    total = len(src)
    for idx, page in enumerate(src, start=1):
        if log_cb: log_cb(f"Watermarking page {idx}/{total}…")
        mat = fitz.Matrix(dpi/72.0, dpi/72.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples).convert("RGBA")
        img = _draw_tiled_watermark(img, wm_text, dpi=dpi, opacity=60)
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
        rect = fitz.Rect(0, 0, pix.width, pix.height)
        new_page = out.new_page(width=rect.width, height=rect.height)
        new_page.insert_image(rect, stream=buf.getvalue())
        if progress_cb:
            progress_cb(idx, total)
    result = out.tobytes()
    out.close(); src.close()
    return result

# ---------------------------
# Conversion pipeline
# ---------------------------
def convert_office_to_pdf_bytes(file_bytes: bytes, in_name: str) -> bytes:
    if not LO_BIN:
        raise RuntimeError("LibreOffice is not available on this host.")
    with tempfile.TemporaryDirectory() as td:
        in_path = pathlib.Path(td) / in_name
        out_dir = pathlib.Path(td) / "out"
        out_dir.mkdir(parents=True, exist_ok=True)
        in_path.write_bytes(file_bytes)
        cmd = [
            LO_BIN,
            "--headless", "--nologo", "--nodefault", "--nolockcheck",
            "--norestore", "--nofirststartwizard",
            "--convert-to", "pdf",
            "--outdir", str(out_dir),
            str(in_path),
        ]
        log_line("Launching LibreOffice conversion…")
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        soft = 0
        pbar = st.session_state["_pbar"]
        while proc.poll() is None:
            if proc.stdout:
                outl = proc.stdout.readline()
                if outl:
                    log_line(outl.strip())
            soft = min(soft + 2, 90)
            pbar.progress(soft)
            time.sleep(0.05)
        if proc.stdout:
            rest = proc.stdout.read() or ""
            for l in rest.splitlines():
                log_line(l)
        if proc.returncode != 0:
            pbar.progress(0)
            raise RuntimeError(f"LibreOffice exit code {proc.returncode}. See logs above.")
        out_candidates = list(out_dir.glob("*.pdf"))
        if not out_candidates:
            raise FileNotFoundError("No PDF produced by LibreOffice.")
        pbar.progress(100)
        log_line("LibreOffice conversion complete.")
        return out_candidates[0].read_bytes()

# ---------------------------
# UI
# ---------------------------
left, right = st.columns([2, 1])
with left:
    uploaded = st.file_uploader(
        "Upload Office/PDF document",
        type=["doc","docx","ppt","pptx","xls","xlsx","odt","odp","ods","pdf"],
        help="Office or OpenDocument files will be converted to PDF, then watermarked."
    )
with right:
    profile = st.radio(
        "Compression profile",
        [
            "High quality (180 dpi, q90)",
            "Balanced (120 dpi, q75)",
            "Smallest (100 dpi, q60)"
        ],
        index=1
    )
    wm_text = st.text_input("Watermark (≤ 15 chars)", value="SLIDESEAL", max_chars=15)
    st.caption("Watermark is tiled, ~8pt, 1-inch spacing, rotated 45°.")

dpi, quality = (120, 75)
if profile.startswith("High"):
    dpi, quality = (180, 90)
elif profile.startswith("Smallest"):
    dpi, quality = (100, 60)

st.subheader("Status")
clear_log()                 # draws fixed-height 4-line box here
pbar = st.progress(0)       # progress bar directly under status box
st.session_state["_pbar"] = pbar

run = st.button("Start conversion")

if run:
    if not uploaded:
        st.error("Please upload a document first.")
        st.stop()
    if len(wm_text) == 0:
        st.error("Please provide a watermark text (1–15 characters).")
        st.stop()
    try:
        name = getattr(uploaded, "name", "upload")
        ext = pathlib.Path(name).suffix.lower()
        if ext == ".pdf":
            log_line("Input is already a PDF. Skipping LibreOffice conversion.")
            pdf_bytes = uploaded.getbuffer().tobytes()
        else:
            log_line(f"Converting '{name}' to PDF via LibreOffice…")
            pdf_bytes = convert_office_to_pdf_bytes(uploaded.getbuffer(), name)

        log_line(f"Applying watermark '{wm_text}' and rebuilding PDF (dpi={dpi}, q={quality})…")
        def page_progress(i, total):
            pct = 10 + int(90 * i / max(1, total))   # reserve 0–10% for LO, then pages to 100%
            pbar.progress(min(pct, 100))

        watermarked = pdf_to_imageonly_pdf_with_watermark(
            pdf_bytes, wm_text, dpi, quality,
            progress_cb=page_progress, log_cb=log_line
        )
        pbar.progress(100)
        log_line("Watermarking complete.")
        out_name = pathlib.Path(name).with_suffix(".pdf").stem + "_sealed.pdf"
        st.success("Done.")
        st.download_button("Download sealed PDF", data=watermarked,
                           file_name=out_name, mime="application/pdf")
    except Exception as e:
        st.error(f"Conversion failed: {e}")