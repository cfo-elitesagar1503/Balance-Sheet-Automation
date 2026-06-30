import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Tally Automation Dashboard", page_icon="📊", layout="wide")

st.title("📊 Tally Automation & Accounting Engine")
st.markdown("Automate your GSTR1, GSTR2B, and Bank Statement accounting entries directly into your Tally Excel Dashboard.")

# Create tabs for different phases of the project
tab1, tab2, tab3, tab4 = st.tabs(["1. Sales (GSTR1)", "2. Purchases (GSTR2B)", "3. Bank Statements", "4. Final Output"])

with tab1:
    st.header("Upload Sales Data")
    gstr1_file = st.file_uploader("Upload GSTR1 (Excel/CSV)", type=["csv", "xlsx", "xls"], key="gstr1")
    form26as_file = st.file_uploader("Upload Form 26AS (PDF)", type=["pdf"], key="26as")
    
    if st.button("Process Sales Data"):
        if gstr1_file is not None and form26as_file is not None:
            st.info("Processing Sales and matching TDS... (Logic to be implemented)")
        else:
            st.warning("Please upload both GSTR1 and Form 26AS files.")

with tab2:
    st.header("Upload Purchase Data")
    gstr2b_file = st.file_uploader("Upload GSTR2B (Excel/CSV)", type=["csv", "xlsx", "xls"], key="gstr2b")
    
    st.markdown("### Client Industry Detection")
    st.markdown("Upload the MOA (Memorandum of Association) so Gemini AI can automatically determine the core business objective and accurately categorize expenses.")
    moa_file = st.file_uploader("Upload MOA (PDF or Image)", type=["pdf", "png", "jpg", "jpeg"], key="moa")
    
    if st.button("Categorize Expenses (AI)"):
        if gstr2b_file is not None and moa_file is not None:
            st.info("Reading MOA using Gemini Vision...")
            import tempfile
            from google import genai
            
            # Save MOA temporarily to pass to Gemini
            with tempfile.NamedTemporaryFile(delete=False, suffix="." + moa_file.name.split('.')[-1]) as tmp_file:
                tmp_file.write(moa_file.read())
                tmp_path = tmp_file.name
                
            try:
                # Initialize Gemini Client
                client = genai.Client()
                uploaded_moa = client.files.upload(file=tmp_path)
                
                moa_prompt = "Read this document. Extract the primary business objective and determine the core industry of the client. Be concise."
                moa_response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[uploaded_moa, moa_prompt]
                )
                detected_industry = moa_response.text
                st.success(f"**Detected Industry:** {detected_industry}")
                
                st.info("Sending Vendor Data to Gemini API for category classification... (Logic to be implemented)")
                # Here we would call process_purchase_data with moa_text=detected_industry
                
            except Exception as e:
                st.error(f"Error reading MOA with AI: {e}")
            finally:
                os.remove(tmp_path)
        else:
            st.warning("Please upload both the GSTR2B file and the MOA file.")

with tab3:
    st.header("Upload Bank Statements")
    bank_file = st.file_uploader("Upload Bank Statement (Excel/CSV)", type=["csv", "xlsx", "xls"], key="bank")
    
    if st.button("Run Bank Reconciliation Engine"):
        if bank_file is not None:
            st.info("Reconciling Sales, Purchases, and processing unknown narrations with AI... (Logic to be implemented)")
        else:
            st.warning("Please upload the Bank Statement.")

with tab4:
    st.header("Generate Final Tally Dashboard")
    st.markdown("Once all phases are completed, you can generate the final Excel file.")
    
    if st.button("Download Final Dashboard Excel"):
        st.success("Tally Dashboard generated successfully! (Logic to be implemented)")

st.sidebar.header("Settings")
# Check if API Key is already configured via secrets or environment variables
secret_api_key = st.secrets.get("GEMINI_API_KEY") if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets else None
env_api_key = os.environ.get("GEMINI_API_KEY")
actual_key = secret_api_key or env_api_key

if actual_key:
    os.environ["GEMINI_API_KEY"] = actual_key
    st.sidebar.success("✅ API Key securely loaded!")
else:
    api_key = st.sidebar.text_input("Gemini API Key", type="password")
    if api_key:
        os.environ["GEMINI_API_KEY"] = api_key
        st.sidebar.success("API Key saved for this session.")
    else:
        st.sidebar.warning("Please enter your Tier 1 Gemini API Key to use AI features.")
