import streamlit as st

# Minimal test app
st.title("Test App")
st.write("If you can see this, the basic app is working!")

# Test basic functionality
if st.button("Test Button"):
    st.success("Button works!")

st.write("App loaded successfully!")
