import streamlit as st
import subprocess, tempfile, pathlib, os

st.set_page_config(page_title="Streamlit + LibreOffice on Render")

st.title("🧪 Streamlit + LibreOffice")
# Version check
try:
    ver = subprocess.check_output(["soffice", "--version"], text=True).strip()
    st.success(f"LibreOffice detected: {ver}")
except Exception as e:
    st.error(f"LibreOffice not found or not runnable: {e}")

st.write("Upload a DOCX/PPTX/ODT and I will try to convert it to PDF using headless LibreOffice.")

uploaded = st.file_uploader("Upload document", type=["doc","docx","ppt","pptx","xls","xlsx","odt","odp","ods"])
if uploaded:
    with tempfile.TemporaryDirectory() as td:
        in_path = pathlib.Path(td) / uploaded.name
        out_dir = pathlib.Path(td) / "out"
        out_dir.mkdir(parents=True, exist_ok=True)
        in_path.write_bytes(uploaded.getbuffer())

        cmd = [
            "soffice",
            "--headless", "--nologo", "--nodefault", "--nolockcheck",
            "--norestore", "--nofirststartwizard",
            "--convert-to", "pdf",
            "--outdir", str(out_dir),
            str(in_path)
        ]
        st.code(" ".join(cmd))
        try:
            subprocess.check_call(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT)
            pdf_name = in_path.with_suffix(".pdf").name
            pdf_path = out_dir / pdf_name
            if pdf_path.exists():
                st.success("Conversion OK.")
                with open(pdf_path, "rb") as f:
                    st.download_button("Download PDF", f, file_name=pdf_name, mime="application/pdf")
            else:
                st.error("Conversion finished but no PDF found.")
        except subprocess.CalledProcessError as e:
            st.error(f"Conversion failed (exit {e.returncode}).")
