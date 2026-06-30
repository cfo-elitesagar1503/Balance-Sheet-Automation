import streamlit as st
import pandas as pd
import os
import tempfile

st.set_page_config(page_title="Tally Automation Dashboard", page_icon="📊", layout="wide")

st.title("📊 Tally Automation & Accounting Engine")
st.markdown("Automate your GSTR1, GSTR2B, and Bank Statement accounting entries directly into your Tally Excel Dashboard.")

os.makedirs("temp_workspace", exist_ok=True)

# Create tabs for different phases of the project
tab1, tab2, tab3, tab4 = st.tabs(["1. Sales (GSTR1)", "2. Purchases (GSTR2B)", "3. Bank Statements", "4. Final Output"])

with tab1:
    st.header("Upload Sales Data")
    gstr1_file = st.file_uploader("Upload GSTR1 (Excel/CSV)", type=["csv", "xlsx", "xls"], key="gstr1")
    form26as_file = st.file_uploader("Upload Form 26AS (PDF)", type=["pdf"], key="26as")
    
    if st.button("Process Sales Data"):
        if gstr1_file is not None:
            st.info("Processing Sales data...")
            from sales_processor import process_sales_data
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_gstr1:
                tmp_gstr1.write(gstr1_file.read())
                
                # Handle optional 26AS file
                tmp_26as_path = None
                if form26as_file is not None:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_26as:
                        tmp_26as.write(form26as_file.read())
                        tmp_26as_path = tmp_26as.name
                
                try:
                    df_final, df_ledgers = process_sales_data(tmp_gstr1.name, None, tmp_26as_path)
                    sales_out = "temp_workspace/sales_output.xlsx"
                    with pd.ExcelWriter(sales_out) as writer:
                        df_final.to_excel(writer, sheet_name="RAW_DATA_MASTER", index=False)
                        df_ledgers.to_excel(writer, sheet_name="LEDGER_GROUP_MAP", index=False)
                    st.session_state['sales_out'] = sales_out
                    st.success("✅ Sales data processed successfully!")
                except Exception as e:
                    st.error(f"Error processing sales: {e}")
        else:
            st.warning("Please upload at least the GSTR1 file.")

with tab2:
    st.header("Upload Purchase Data")
    gstr2b_file = st.file_uploader("Upload GSTR2B (Excel/CSV)", type=["csv", "xlsx", "xls"], key="gstr2b")
    
    st.markdown("### Client Industry Detection")
    st.markdown("Upload the MOA (Memorandum of Association) so Gemini AI can automatically determine the core business objective and accurately categorize expenses.")
    moa_file = st.file_uploader("Upload MOA (PDF or Image)", type=["pdf", "png", "jpg", "jpeg"], key="moa")
    
    if st.button("Categorize Expenses (AI)"):
        if gstr2b_file is not None and moa_file is not None:
            st.info("Reading MOA using Gemini Vision...")
            from google import genai
            from purchase_processor import process_purchase_data
            
            with tempfile.NamedTemporaryFile(delete=False, suffix="." + moa_file.name.split('.')[-1]) as tmp_file:
                tmp_file.write(moa_file.read())
                tmp_path = tmp_file.name
                
            try:
                client = genai.Client()
                uploaded_moa = client.files.upload(file=tmp_path)
                
                moa_prompt = "Read this document. Extract the primary business objective and determine the core industry of the client. Be concise."
                moa_response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[uploaded_moa, moa_prompt]
                )
                detected_industry = moa_response.text
                st.success(f"**Detected Industry:** {detected_industry}")
                
                st.info("Processing Vendor Data and classifying categories...")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_gstr2b:
                    tmp_gstr2b.write(gstr2b_file.read())
                    df_final, df_ledgers = process_purchase_data(tmp_gstr2b.name, detected_industry)
                    purch_out = "temp_workspace/purchases_output.xlsx"
                    with pd.ExcelWriter(purch_out) as writer:
                        df_final.to_excel(writer, sheet_name="RAW_DATA_MASTER", index=False)
                        df_ledgers.to_excel(writer, sheet_name="LEDGER_GROUP_MAP", index=False)
                    st.session_state['purch_out'] = purch_out
                    st.success("✅ Purchase data processed successfully!")
                    
            except Exception as e:
                st.error(f"Error reading MOA with AI or processing purchases: {e}")
            finally:
                os.remove(tmp_path)
        else:
            st.warning("Please upload both the GSTR2B file and the MOA file.")

with tab3:
    st.header("Upload Bank Statements")
    bank_file = st.file_uploader("Upload Bank Statement (Excel/CSV)", type=["csv", "xlsx", "xls"], key="bank")
    
    if st.button("Run Bank Reconciliation Engine"):
        if bank_file is not None:
            if 'sales_out' not in st.session_state or 'purch_out' not in st.session_state:
                st.warning("Please process Sales and Purchases first!")
            else:
                st.info("Reconciling Sales, Purchases, and processing unknown narrations with AI...")
                from bank_processor import process_bank_statement
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_bank:
                    tmp_bank.write(bank_file.read())
                    try:
                        bank_df = process_bank_statement(tmp_bank.name, st.session_state['sales_out'], st.session_state['purch_out'])
                        bank_out = "temp_workspace/bank_output.xlsx"
                        bank_df.to_excel(bank_out, sheet_name="Bank", index=False)
                        st.session_state['bank_out'] = bank_out
                        st.success("✅ Bank Reconciliation Engine completed successfully!")
                    except Exception as e:
                        st.error(f"Error processing bank data: {e}")
        else:
            st.warning("Please upload the Bank Statement.")

with tab4:
    st.header("Generate Final Tally Dashboard")
    st.markdown("Once all phases are completed, you can generate the final Excel file.")
    
    if st.button("Generate Final Dashboard Excel"):
        if 'sales_out' in st.session_state and 'purch_out' in st.session_state and 'bank_out' in st.session_state:
            from combine_to_dashboard import create_dashboard
            final_out = "temp_workspace/Final_Tally_Dashboard.xlsx"
            try:
                create_dashboard(st.session_state['sales_out'], st.session_state['purch_out'], st.session_state['bank_out'], final_out)
                st.success("✅ Tally Dashboard generated successfully!")
                
                with open(final_out, "rb") as file:
                    st.download_button(
                        label="⬇️ Download Final Dashboard Excel",
                        data=file,
                        file_name="Final_Tally_Dashboard.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Error generating dashboard: {e}")
        else:
            st.warning("Please complete processing in Sales, Purchases, and Bank tabs first.")

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
