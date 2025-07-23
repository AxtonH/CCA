import streamlit as st
import sys
import os

# Basic health check
st.write("Python version:", sys.version)
st.write("Working directory:", os.getcwd())
st.write("Files in directory:", os.listdir("."))

# Minimal test app
st.title("Test App")
st.write("If you can see this, the basic app is working!")

# Test basic functionality
if st.button("Test Button"):
    st.success("Button works!")

st.write("App loaded successfully!")
