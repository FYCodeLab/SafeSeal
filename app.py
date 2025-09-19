import streamlit as st
import subprocess, tempfile, pathlib, time, os, shutil

st.set_page_config(page_title="SafeSeal · Streamlit + LibreOffice", layout="centered")

# --- Small CSS for compact status window text
st.markdown("""
<style>
.small-text {font-size: 12px; line-height: 1.2;}
.status-box {
  border: 1px solid #ddd; border-radius: 6px; padding: 8px 10px; background: #fafafa;
  max-height: 220px; overflow-y: auto;
}
code, pre {font-size: 11px;}
</style>
""", unsafe_allow_html=True)

st.title("SafeSeal · Document → PDF (LibreOffice)")

# Quick environment probe
colA, colB = st.columns(2)
with colA:
    try:
        ver = subprocess.check_output(["soffice", "--version"], text=True).strip()
        st.success(f"LibreOffice: {ver}")
    except Exception as e:
        st.error(f"LibreOffice not detected: {e}")
with colB:
    py = subprocess.check_output(["python", "--version"], text=True).strip()
    st.info(py)

st.write("Upload a document and I will convert it to PDF using headless LibreOffice.")

uploaded = st.file_uploader(
    "Choose a file",
    type=["doc","docx","ppt","pptx","xls","xlsx","odt","odp","ods"],
    help="Microsoft Office, OpenDocument formats are supported."
)

def convert_with_libreoffice(input_bytes: bytes, filename: str) -> pathlib.Path:
    """Run headless LibreOffice conversion. Returns path to produced PDF."""
    with tempfile.TemporaryDirectory() as td:
        td_path = pathlib.Path(td)
        in_path = td_path / filename
        out_dir = td_path / "out"
        out_dir.mkdir(parents=True, exist_ok=True)

        in_path.write_bytes(input_bytes)

        # Build command
        cmd = [
            "soffice",
            "--headless", "--nologo", "--nodefault", "--nolockcheck",
            "--norestore", "--nofirststartwizard",
            "--convert-to", "pdf",
            "--outdir", str(out_dir),
            str(in_path),
        ]

        # Launch process without blocking to allow UI updates
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True
        )

        # Stream logs and animate progress
        log_lines = []
        progress = 0
        last_bump = time.time()

        status_box = st.empty()      # the log window
        pbar = st.progress(0)        # the progress bar

        def render_logs():
            # Render last ~200 lines, compact small font
            st_html = "<div class='status-box small-text'>" + "<br/>".join(
                [l.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;") for l in log_lines[-200:]]
            ) + "</div>"
            status_box.markdown(st_html, unsafe_allow_html=True)

        # Read output line-by-line while faking smooth progress
        while proc.poll() is None:
            # Non-blocking read of any available output
            if proc.stdout:
                line = proc.stdout.readline()
                if line:
                    log_lines.append(line.rstrip())
                    render_logs()

            # Bump the progress slowly while the process is alive
            now = time.time()
            if now - last_bump > 0.2:
                # Ease toward 90% while running
                progress = min(progress + 2, 90)
                pbar.progress(progress)
                last_bump = now

            time.sleep(0.05)

        # Drain remaining output
        if proc.stdout:
            rest = proc.stdout.read() or ""
            if rest:
                log_lines.extend([l for l in rest.splitlines()])
                render_logs()

        rc = proc.returncode
        if rc != 0:
            pbar.progress(0)
            raise RuntimeError(f"LibreOffice exited with code {rc}. See logs above.")

        # Complete progress
        pbar.progress(100)

        # Figure out output name
        pdf_name = pathlib.Path(filename).with_suffix(".pdf").name
        pdf_path = out_dir / pdf_name
        if not pdf_path.exists():
            # Some formats rename oddly; fallback: find first PDF in out_dir
            candidates = list(out_dir.glob("*.pdf"))
            if not candidates:
                raise FileNotFoundError("Conversion finished but no PDF was produced.")
            pdf_path = candidates[0]

        # Move PDF to a persistent temp so it survives context exit
        final_dir = pathlib.Path(tempfile.mkdtemp())
        final_pdf = final_dir / pdf_path.name
        shutil.copy2(pdf_path, final_pdf)
        return final_pdf

if uploaded:
    st.subheader("Status")
    try:
        pdf_path = convert_with_libreoffice(uploaded.getbuffer(), uploaded.name)
        st.success("Conversion complete.")
        with open(pdf_path, "rb") as f:
            st.download_button("Download PDF", f, file_name=pdf_path.name, mime="application/pdf")
        st.caption(f"Saved temporary file: {pdf_path}")
    except Exception as e:
        st.error(f"Conversion failed: {e}")
        st.stop()

st.caption("Tip: first request after an idle period may be slow due to platform cold start.")