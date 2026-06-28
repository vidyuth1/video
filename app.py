import streamlit as st
import cv2
import numpy as np
import tempfile
import os

st.title("🍫 Chocolate Mould Video Inspector")
st.write("Upload your conveyor belt video to strip out empty belt frames automatically.")

# 1. Sidebar controls for tuning settings dynamically
st.sidebar.header("Tuning Settings")
history = st.sidebar.slider("Background History (Frames)", 50, 500, 300, help="How many frames back the system looks to define 'empty belt'.")
var_threshold = st.sidebar.slider("Subtractor Threshold", 16, 100, 60, help="Lower values detect subtle changes; higher values ignore minor lighting shifts.")
pixel_threshold = st.sidebar.slider("Min Trigger Pixels", 100, 2000, 600, help="How many white pixels must change to flag a mould.")

uploaded_file = st.file_uploader("Upload Conveyor Video (MP4 format recommended)", type=["mp4", "avi", "mov"])

if uploaded_file is not None:
    # Streamlit uploads are bytes-like objects, so we write to a temporary file for OpenCV to read
    tfile = tempfile.NamedTemporaryFile(delete=False)
    tfile.write(uploaded_file.read())
    
    cap = cv2.VideoCapture(tfile.name)
    
    width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    fps = cap.get(cv2.CAP_PROP_FPS)
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    
    # Define temp output file path
    output_path = os.path.join(tempfile.gettempdir(), "moulds_only.mp4")
    
    # Use H264 codec for web-compatible playback downloads
    fourcc = cv2.VideoWriter_fourcc(*'H264')
    if not fourcc: # Fallback if H264 isn't locally available on the container
        fourcc = cv2.VideoWriter_fourcc(*'mp4v')
        
    out = cv2.VideoWriter(output_path, fourcc, fps, (width, height))
    
    # Initialize background subtractor with sidebar variables
    bg_subtractor = cv2.createBackgroundSubtractorMOG2(history=history, varThreshold=var_threshold, detectShadows=False)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    frames_processed = 0
    saved_frames = 0
    
    while cap.isOpened():
        ret, frame = cap.read()
        if not ret:
            break
            
        # Regional strip-wire logic
        mid_y = int(height / 2)
        roi = frame[mid_y - 80 : mid_y + 80, 0 : width]
        
        mask = bg_subtractor.apply(roi)
        _, mask = cv2.threshold(mask, 200, 255, cv2.THRESH_BINARY)
        changed_pixels = cv2.countNonZero(mask)
        
        if changed_pixels > pixel_threshold:
            out.write(frame)
            saved_frames += 1
            
        frames_processed += 1
        # Update progress bar occasionally to save processing overhead
        if frames_processed % 30 == 0:
            progress = frames_processed / total_frames
            progress_bar.progress(progress)
            status_text.text(f"Processing frame {frames_processed}/{total_frames}...")
            
    cap.release()
    out.release()
    
    progress_bar.progress(1.0)
    status_text.text("Processing Complete!")
    
    st.success(f"Done! Kept {saved_frames} frames out of {total_frames} total frames.")
    
    # Provide direct download link for the completed video
    with open(output_path, "rb") as file:
        st.download_button(
            label="📥 Download Trimmed Video",
            data=file,
            file_name="moulds_only_final.mp4",
            mime="video/mp4"
        )
        
    # Cleanup temp files
    os.unlink(tfile.name)
