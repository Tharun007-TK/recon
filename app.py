import streamlit as st
import pandas as pd
from engine import reconcile, generate_supplier_defaulters_list

st.set_page_config(
    page_title="Recko: GST vs Tally Reconciliation Engine",
    page_icon="📊",
    layout="wide"
)

def main():
    # Title and Introduction
    st.title("📊 Recko: GST vs Tally Reconciliation Engine")
    st.markdown("""
    Welcome to **Recko**. This tool reconciles your Tally Purchase Register against the GSTR-2B report.
    Upload your files below to identify matched, mismatched, and missing invoices.
    """)
    
    st.divider()

    # File Uploaders
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Upload Tally Export")
        tally_file = st.file_uploader(
            "Upload Tally Purchase Register (Excel format)", 
            type=['xlsx', 'xls'], 
            key='tally'
        )
        st.caption("Ensure it contains Voucher No/Bill No, GSTIN, Taxable Value, and GST Amount.")

    with col2:
        st.subheader("2. Upload GSTR-2B Export")
        gstr2b_file = st.file_uploader(
            "Upload GST Portal GSTR-2B (Excel format)", 
            type=['xlsx', 'xls'], 
            key='gstr2b'
        )
        st.caption("Ensure it contains the B2B table with Invoice Number, GSTIN of supplier, and Tax amounts.")

    st.divider()

    if st.button("🚀 Run Reconciliation", type="primary", use_container_width=True):
        if tally_file is None or gstr2b_file is None:
            st.warning("⚠️ Please upload both Tally and GSTR-2B files to proceed.")
        else:
            try:
                with st.spinner("Processing and reconciling data. This may take a moment..."):
                    # Call the reconciliation engine
                    summary_metrics, excel_file, warnings, missing_in_gst_df = reconcile(tally_file, gstr2b_file)
                    
                    # Store results in session state so they persist after clicking a download button
                    st.session_state['recon_results'] = {
                        'summary_metrics': summary_metrics,
                        'excel_file': excel_file,
                        'warnings': warnings,
                        'missing_in_gst_df': missing_in_gst_df
                    }
            except Exception as e:
                # Handle and display errors cleanly
                import traceback
                st.error(f"❌ An error occurred during reconciliation: {str(e)}")
                st.text(traceback.format_exc())
                st.markdown("Please verify your uploaded files format and try again.")

    # Render results if they exist in session state
    if 'recon_results' in st.session_state:
        res = st.session_state['recon_results']
        summary_metrics = res['summary_metrics']
        excel_file = res['excel_file']
        warnings = res['warnings']
        missing_in_gst_df = res['missing_in_gst_df']
        
        st.success("✅ Reconciliation completed successfully!")
        
        if warnings:
            for w in warnings:
                st.warning(w)
        
        # Display Dashboard Metrics
        st.header("Dashboard Summary")
        
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Records", summary_metrics.get('Total Records', 0))
        m2.metric("Matched ✅", summary_metrics.get('Matched', 0))
        m3.metric("Mismatched ⚠️", summary_metrics.get('Mismatched', 0))
        m4.metric("Missing in GST ❌", summary_metrics.get('Missing in GST', 0))
        m5.metric("Missing in Tally ❌", summary_metrics.get('Missing in Tally', 0))
        
        st.divider()
        
        # Downloadable Report
        st.header("Download Detailed Report")
        st.markdown("The detailed Excel report contains separate sheets for Matched, Mismatched, and Missing records with explicit reasons for any discrepancies.")
        
        st.download_button(
            label="📥 Download Reconciliation Report (Excel)",
            data=excel_file,
            file_name="Reconciliation_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.divider()
        st.header("📧 Supplier Follow-Up Tools")

        # Check if there are any missing in GST records
        if summary_metrics.get('Missing in GST', 0) > 0:
            st.warning(f"⚠️ **{summary_metrics['Missing in GST']} invoices** are missing in GST Portal. These suppliers haven't filed their GSTR-1!")
            
            # Generate and download Supplier Defaulters List
            try:
                defaulters_file = generate_supplier_defaulters_list(missing_in_gst_df)
                
                if defaulters_file:
                    st.download_button(
                        label="📥 Download Supplier Defaulters List (Excel)",
                        data=defaulters_file,
                        file_name=f"Supplier_Defaulters_List_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="secondary",
                        use_container_width=True
                    )
                    
                    st.info("""
                    **💡 How to use this list:**
                    1. Open the Excel file
                    2. Filter by Supplier Name
                    3. Send email/WhatsApp to each supplier with their pending invoices
                    4. Ask them to file GSTR-1 immediately
                    5. Once filed, their invoices will appear in your GSTR-2B
                    """)
            except Exception as e:
                st.error(f"Error generating defaulters list: {str(e)}")
        else:
            st.success("✅ All suppliers have filed their GSTR-1! No defaulters found.")

if __name__ == "__main__":
    main()
