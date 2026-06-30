import streamlit as st
import pandas as pd
import os

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
            
            gstr1_path = "temp_workspace/gstr1.xlsx"
            with open(gstr1_path, "wb") as f:
                f.write(gstr1_file.getvalue())
                
            form26as_path = None
            if form26as_file is not None:
                form26as_path = "temp_workspace/form26as.pdf"
                with open(form26as_path, "wb") as f:
                    f.write(form26as_file.getvalue())
            
            try:
                df_final, df_ledgers = process_sales_data(gstr1_path, None, form26as_path)
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

    if 'sales_out' in st.session_state and os.path.exists(st.session_state['sales_out']):
        st.markdown("---")
        with open(st.session_state['sales_out'], "rb") as file:
            st.download_button("⬇️ Download Processed Sales Excel", data=file, file_name="Processed_Sales.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
            
            moa_path = "temp_workspace/moa_file." + moa_file.name.split('.')[-1]
            with open(moa_path, "wb") as f:
                f.write(moa_file.getvalue())
                
            try:
                client = genai.Client()
                uploaded_moa = client.files.upload(file=moa_path)
                
                moa_prompt = "Read this document. Extract the primary business objective and determine the core industry of the client. Be concise."
                moa_response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[uploaded_moa, moa_prompt]
                )
                detected_industry = moa_response.text
                st.success(f"**Detected Industry:** {detected_industry}")
                
                st.info("Processing Vendor Data and classifying categories...")
                gstr2b_path = "temp_workspace/gstr2b.xlsx"
                with open(gstr2b_path, "wb") as f:
                    f.write(gstr2b_file.getvalue())
                
                df_final, df_ledgers = process_purchase_data(gstr2b_path, detected_industry)
                purch_out = "temp_workspace/purchases_output.xlsx"
                with pd.ExcelWriter(purch_out) as writer:
                    df_final.to_excel(writer, sheet_name="RAW_DATA_MASTER", index=False)
                    df_ledgers.to_excel(writer, sheet_name="LEDGER_GROUP_MAP", index=False)
                st.session_state['purch_out'] = purch_out
                st.success("✅ Purchase data processed successfully!")
                    
            except Exception as e:
                st.error(f"Error reading MOA with AI or processing purchases: {e}")
        else:
            st.warning("Please upload both the GSTR2B file and the MOA file.")

    if 'purch_out' in st.session_state and os.path.exists(st.session_state['purch_out']):
        st.markdown("---")
        with open(st.session_state['purch_out'], "rb") as file:
            st.download_button("⬇️ Download Processed Purchases Excel", data=file, file_name="Processed_Purchases.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab3:
    st.header("Upload Bank Statements")
    st.markdown("### Pre-processed Data (Optional)")
    st.markdown("If you already processed Sales and Purchases in a previous session, upload their Excel files here to save AI credits. Otherwise, it will use the data from Tab 1 & 2.")
    
    col1, col2 = st.columns(2)
    with col1:
        processed_sales_file = st.file_uploader("Upload Processed Sales", type=["xlsx"], key="psales")
    with col2:
        processed_purch_file = st.file_uploader("Upload Processed Purchases", type=["xlsx"], key="ppurch")

    st.markdown("### Bank Statement")
    bank_file = st.file_uploader("Upload Bank Statement (Excel/CSV)", type=["csv", "xlsx", "xls"], key="bank")
    
    if st.button("Run Bank Reconciliation Engine"):
        if bank_file is not None:
            sales_path = None
            purch_path = None
            
            if processed_sales_file:
                sales_path = "temp_workspace/uploaded_sales_out.xlsx"
                with open(sales_path, "wb") as f:
                    f.write(processed_sales_file.getvalue())
                st.session_state['sales_out'] = sales_path
            elif 'sales_out' in st.session_state:
                sales_path = st.session_state['sales_out']
                
            if processed_purch_file:
                purch_path = "temp_workspace/uploaded_purch_out.xlsx"
                with open(purch_path, "wb") as f:
                    f.write(processed_purch_file.getvalue())
                st.session_state['purch_out'] = purch_path
            elif 'purch_out' in st.session_state:
                purch_path = st.session_state['purch_out']
                
            if not sales_path or not purch_path:
                st.warning("Please either upload the Processed Sales/Purchases files above OR process them in Tab 1 & 2 first!")
            else:
                st.info("Reconciling Sales, Purchases, and processing unknown narrations with AI...")
                from bank_processor import process_bank_statement
                
                bank_path = "temp_workspace/bank_upload.xlsx"
                with open(bank_path, "wb") as f:
                    f.write(bank_file.getvalue())
                
                try:
                    bank_df = process_bank_statement(bank_path, sales_path, purch_path)
                    bank_out = "temp_workspace/bank_output.xlsx"
                    bank_df.to_excel(bank_out, sheet_name="Bank", index=False)
                    st.session_state['bank_out'] = bank_out
                    st.success("✅ Bank Reconciliation Engine completed successfully!")
                except Exception as e:
                    st.error(f"Error processing bank data: {e}")
        else:
            st.warning("Please upload the Bank Statement.")

    if 'bank_out' in st.session_state and os.path.exists(st.session_state['bank_out']):
        st.markdown("---")
        with open(st.session_state['bank_out'], "rb") as file:
            st.download_button("⬇️ Download Analysed Bank Excel", data=file, file_name="Analysed_Bank.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
