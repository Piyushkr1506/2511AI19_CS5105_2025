# streamlit_app.py
import streamlit as st
import tempfile
import os
import shutil

from seating_allocator import SeatingAllocator
from logger_setup import setup_logging


def close_logger(logger):
    """Release file handles so TemporaryDirectory can clean up on Windows."""
    if logger is None:
        return
    # Copy the list so we can modify logger.handlers while iterating
    for h in list(logger.handlers):
        try:
            h.flush()
        except Exception:
            pass
        try:
            h.close()
        except Exception:
            pass
        logger.removeHandler(h)


def run_allocation(uploaded_file, buffer, density):
    # This temp dir (and everything inside) will be deleted automatically
    with tempfile.TemporaryDirectory() as tmpdir:
        # Save uploaded Excel to a temp path
        excel_path = os.path.join(tmpdir, uploaded_file.name)
        with open(excel_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Create output folder FIRST
        outdir = os.path.join(tmpdir, "output")
        os.makedirs(outdir, exist_ok=True)

        # Now it's safe to create log files inside outdir
        logger = setup_logging(
            logfile=os.path.join(outdir, "seating.log"),
            errorfile=os.path.join(outdir, "errors.txt"),
        )

        # Run allocation pipeline
        alloc = SeatingAllocator(
            input_file=excel_path,
            buffer=buffer,
            density=density,
            outdir=outdir,
            logger=logger,
        )
        alloc.load_inputs()
        alloc.allocate_all_days()
        alloc.write_outputs()

        # Generate attendance PDFs
        photos_dir = "photos"  # keep this dir next to the app
        no_image_icon = os.path.join(photos_dir, "no_image_available.jpg")
        alloc.generate_attendance_pdfs(photos_dir, no_image_icon)

        # Zip the entire output folder
        zip_base = os.path.join(tmpdir, "attendance_output")
        shutil.make_archive(zip_base, "zip", outdir)
        zip_path = zip_base + ".zip"

        # Read the zip into memory BEFORE TemporaryDirectory is cleaned up
        with open(zip_path, "rb") as f:
            zip_bytes = f.read()

        # Very important on Windows: release errors.txt / seating.log
        close_logger(logger)

        # Return bytes, not a path inside the soon-to-be-deleted temp dir
        return zip_bytes


st.title("IIT Patna Exam Seating & Attendance sheet Generator")

uploaded = st.file_uploader("Upload input Excel file", type=["xlsx"])
buffer = st.number_input("Buffer seats (per room)", min_value=0, max_value=50, value=0)
density = st.selectbox("Seating density", ["Dense", "Sparse"])

if st.button("Generate seating + attendance") and uploaded:
    with st.spinner("Running allocation..."):
        try:
            zip_bytes = run_allocation(uploaded, buffer, density)
            st.download_button(
                "Download output ZIP",
                data=zip_bytes,
                file_name="seating_attendance_output.zip",
                mime="application/zip",
            )
        except Exception as e:
            st.error(f"Error: {e}")
