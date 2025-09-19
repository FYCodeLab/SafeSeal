# app.py — SafeSeal v5.2 · Streamlit + LibreOffice + Watermark
# - Status box: dark grey, green Courier, very small font
# - Status box: hidden until user clicks "Start conversion"
# - Fixed 4-line height with auto-scroll to newest line

import io, os, shutil, subprocess, tempfile, time, pathlib
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image, ImageDraw, ImageFont
import fitz  # PyMuPDF

# ---------------------------
# Page setup
# ---------------------------
st.set_page_config(page_title="SafeSeal v5.2 · Streamlit + LibreOffice + Watermark",
                   layout="centered")
st.subheader("SafeSeal v5.2 · Streamlit + LibreOffice + Watermark")

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
# Status rendering (inline styles so they work inside iframe)
# ---------------------------
def _html_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# Reserve the spot under the header; we will draw here later
st.subheader("Status")
_status_placeholder = st.empty()       # status box appears here after click
_pbar_placeholder   = st.empty()       # progress bar appears here after click

# Session state defaults
st.session_state.setdefault("_logbuf", "")
st.session_state.setdefault("_show_status", False)  # hidden by default until user clicks
st.session_state.setdefault("_pbar", None)

def _render_status_box():
    """Render current log buffer into a fixed 4-line box (dark grey, green Courier) with auto-scroll."""
    buf = st.session_state.get("_logbuf", "")
    # Very small font (11px), line-height 1.2 → ~4 lines tall + padding
    html = f"""
<div id="status-box"
     style="background:#3a3a3a; color:#00ff00;
            border:1px solid #5a5a5a; border-radius:6px;
            padding:8px 10px;
            font-family:'Courier New', Courier, monospace;
            font-size:11px; line-height:1.2;
            white-space:pre-wrap;
            height:calc(1.2em * 4 + 16px);
            overflow-y:auto; overflow-x:hidden;">
{_html_escape(buf)}
</div>
<script>
  const el = document.getElementById('status-box');
  if (el) {{
    el.scrollTop = el.scrollHeight;  // jump to latest line
  }}
</script>
"""
    # Render in the reserved spot (use iframe so JS runs)
    with _status_placeholder.container():
        components.html(html, height=110, scrolling=False)

def _clear_log():
    st.session_state["_logbuf"] = ""
    if st.session_state["_show_status"]:
        _render_status_box()

def _log_line(msg: str):
    st.session_state["_logbuf"] = st.session_state.get("_logbuf", "") + msg.rstrip() + "\n"
    if st.session_state["_show_status"]:
        _render_status_box()

# If still hidden, show nothing under "Status"; when shown, draw box + progress bar
if st.session_state["_show_status"]:
    _render_status_box()
    if st.session_state["_pbar"] is None:
        st.session_state["_pbar"] = _pbar_placeholder.progress(0)

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
    spacing_px = int(dpi)  # 1 inch spacing
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
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("RGBA")
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
        _log_line("Launching LibreOffice conversion…")
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        soft = 0
        while proc.poll() is None:
            if proc.stdout:
                outl = proc.stdout.readline()
                if outl:
                    _log_line(outl.strip())
            soft = min(soft + 2, 90)     # soft progress toward 90% while LO runs
            if st.session_state["_pbar"]:
                st.session_state["_pbar"].progress(soft)
            time.sleep(0.05)
        if proc.stdout:
            rest = proc.stdout.read() or ""
            for l in rest.splitlines():
                _log_line(l)
        if proc.returncode != 0:
            if st.session_state["_pbar"]:
                st.session_state["_pbar"].progress(0)
            raise RuntimeError(f"LibreOffice exit code {proc.returncode}. See logs above.")
        out_candidates = list(out_dir.glob("*.pdf"))
        if not out_candidates:
            raise FileNotFoundError("No PDF produced by LibreOffice.")
        if st.session_state["_pbar"]:
            st.session_state["_pbar"].progress(100)
        _log_line("LibreOffice conversion complete.")
        return out_candidates[0].read_bytes()

# ---------------------------
# UI (inputs)
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

# ---------------------------
# Run button (reveals status box on click)
# ---------------------------
run = st.button("Start conversion")

if run:
    if not uploaded:
        st.error("Please upload a document first.")
        st.stop()
    if len(wm_text) == 0:
        st.error("Please provide a watermark text (1–15 characters).")
        st.stop()

    # Reveal status box + progress bar now
    st.session_state["_show_status"] = True
    _render_status_box()
    if st.session_state["_pbar"] is None:
        st.session_state["_pbar"] = _pbar_placeholder.progress(0)

    try:
        name = getattr(uploaded, "name", "upload")
        ext = pathlib.Path(name).suffix.lower()
        if ext == ".pdf":
            _log_line("Input is already a PDF. Skipping LibreOffice conversion.")
            pdf_bytes = uploaded.getbuffer().tobytes()
        else:
            _log_line(f"Converting '{name}' to PDF via LibreOffice…")
            pdf_bytes = convert_office_to_pdf_bytes(uploaded.getbuffer(), name)

        _log_line(f"Applying watermark '{wm_text}' and rebuilding PDF (dpi={dpi}, q={quality})…")
        def page_progress(i, total):
            pct = 10 + int(90 * i / max(1, total))   # reserve 0–10% for LO, then pages to 100%
            if st.session_state["_pbar"]:
                st.session_state["_pbar"].progress(min(pct, 100))

        watermarked = pdf_to_imageonly_pdf_with_watermark(
            pdf_bytes, wm_text, dpi, quality,
            progress_cb=page_progress, log_cb=_log_line
        )
        if st.session_state["_pbar"]:
            st.session_state["_pbar"].progress(100)
        _log_line("Watermarking complete.")
        out_name = pathlib.Path(name).with_suffix(".pdf").stem + "_sealed.pdf"
        st.success("Done.")
        st.download_button("Download sealed PDF", data=watermarked,
                           file_name=out_name, mime="application/pdf")
    except Exception as e:
        st.error(f"Conversion failed: {e}")