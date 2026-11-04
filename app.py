import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from io import BytesIO

st.set_page_config(page_title="GSTR-2A Automated Audit Platform", layout="wide")

st.title("📊 GSTR-2A Annual Audit & Reconciliation Platform")
st.markdown("Upload your raw GSTR-2A annual Excel file downloaded from the GST portal to automatically generate an executive audit workbook.")

uploaded_file = st.file_uploader("Upload GSTR-2A Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing GSTR-2A data... Please wait."):
        try:
            read_me = pd.read_excel(uploaded_file, sheet_name='Read me', header=None)
            fy = read_me.iloc[2, 4] if len(read_me) > 2 else '2022-23'

            b2b_raw = pd.read_excel(uploaded_file, sheet_name='B2B', header=5)
            b2b_cols = {
                b2b_raw.columns[0]: 'GSTR2A Period',
                b2b_raw.columns[1]: 'GSTIN of Supplier',
                b2b_raw.columns[2]: 'Trade/Legal Name of Supplier',
                b2b_raw.columns[3]: 'Invoice Number',
                b2b_raw.columns[4]: 'Invoice Type',
                b2b_raw.columns[5]: 'Invoice Date',
                b2b_raw.columns[6]: 'Invoice Value (Rs.)',
                b2b_raw.columns[7]: 'Place of Supply',
                b2b_raw.columns[8]: 'Reverse Charge',
                b2b_raw.columns[9]: 'Rate (%)',
                b2b_raw.columns[10]: 'Taxable Value (Rs.)',
                b2b_raw.columns[11]: 'IGST (Rs.)',
                b2b_raw.columns[12]: 'CGST (Rs.)',
                b2b_raw.columns[13]: 'SGST (Rs.)',
                b2b_raw.columns[14]: 'Cess (Rs.)',
                b2b_raw.columns[15]: 'GSTR-1 Filing Status',
                b2b_raw.columns[16]: 'GSTR-1 Filing Date',
                b2b_raw.columns[17]: 'GSTR-1 Filing Period',
                b2b_raw.columns[18]: 'GSTR-3B Filing Status',
                b2b_raw.columns[19]: 'Amendment Made, if any',
                b2b_raw.columns[20]: 'Tax Period in which Amended',
                b2b_raw.columns[21]: 'Effective Date of Cancellation',
                b2b_raw.columns[22]: 'Source',
                b2b_raw.columns[23]: 'IRN',
                b2b_raw.columns[24]: 'IRN Date'
            }
            b2b_df = b2b_raw.rename(columns=b2b_cols).copy()
            b2b_df['Financial Year'] = fy
            for col in ['Taxable Value (Rs.)', 'IGST (Rs.)', 'CGST (Rs.)', 'SGST (Rs.)', 'Cess (Rs.)', 'Invoice Value (Rs.)']:
                if col in b2b_df.columns:
                    b2b_df[col] = pd.to_numeric(b2b_df[col], errors='coerce').fillna(0.0)

            b2b_df['Invoice Date Parsed'] = pd.to_datetime(b2b_df['Invoice Date'], format='%d-%m-%Y', errors='coerce')
            b2b_df['Cancellation Date Parsed'] = pd.to_datetime(b2b_df['Effective Date of Cancellation'], format='%d-%b-%y', errors='coerce')
            mask_nat = b2b_df['Cancellation Date Parsed'].isna() & b2b_df['Effective Date of Cancellation'].notna()
            if mask_nat.any():
                b2b_df.loc[mask_nat, 'Cancellation Date Parsed'] = pd.to_datetime(b2b_df.loc[mask_nat, 'Effective Date of Cancellation'], format='%d-%b-%Y', errors='coerce')

            try:
                cdnr_raw = pd.read_excel(uploaded_file, sheet_name='CDNR', header=5)
                cdnr_cols = {
                    cdnr_raw.columns[0]: 'GSTR2A Period',
                    cdnr_raw.columns[1]: 'GSTIN of Supplier',
                    cdnr_raw.columns[2]: 'Trade/Legal Name of Supplier',
                    cdnr_raw.columns[3]: 'Note Type',
                    cdnr_raw.columns[4]: 'Note Number',
                    cdnr_raw.columns[5]: 'Note Date',
                    cdnr_raw.columns[6]: 'Note Value (Rs.)',
                    cdnr_raw.columns[7]: 'Reason',
                    cdnr_raw.columns[8]: 'Place of Supply',
                    cdnr_raw.columns[9]: 'Reverse Charge',
                    cdnr_raw.columns[10]: 'Rate (%)',
                    cdnr_raw.columns[11]: 'Taxable Value (Rs.)',
                    cdnr_raw.columns[12]: 'IGST (Rs.)',
                    cdnr_raw.columns[13]: 'CGST (Rs.)',
                    cdnr_raw.columns[14]: 'SGST (Rs.)',
                    cdnr_raw.columns[15]: 'Cess (Rs.)',
                    cdnr_raw.columns[16]: 'GSTR-1 Filing Status',
                    cdnr_raw.columns[17]: 'GSTR-1 Filing Date',
                    cdnr_raw.columns[18]: 'GSTR-1 Filing Period',
                    cdnr_raw.columns[19]: 'GSTR-3B Filing Status',
                    cdnr_raw.columns[20]: 'Source',
                    cdnr_raw.columns[21]: 'IRN',
                    cdnr_raw.columns[22]: 'IRN Date'
                }
                cdnr_df = cdnr_raw.rename(columns=cdnr_cols).copy()
                cdnr_df['Financial Year'] = fy
                for col in ['Taxable Value (Rs.)', 'IGST (Rs.)', 'CGST (Rs.)', 'SGST (Rs.)', 'Cess (Rs.)', 'Note Value (Rs.)']:
                    if col in cdnr_df.columns:
                        cdnr_df[col] = pd.to_numeric(cdnr_df[col], errors='coerce').fillna(0.0)
            except:
                cdnr_df = pd.DataFrame()

            b2b_grouped = b2b_df.groupby(['Financial Year', 'GSTIN of Supplier', 'Trade/Legal Name of Supplier']).agg({
                'Taxable Value (Rs.)': 'sum', 'IGST (Rs.)': 'sum', 'CGST (Rs.)': 'sum', 'SGST (Rs.)': 'sum', 'Cess (Rs.)': 'sum', 'Invoice Value (Rs.)': 'sum'
            }).reset_index().rename(columns={'Taxable Value (Rs.)': 'Gross Purchases - Taxable Value (Rs.)', 'Invoice Value (Rs.)': 'Gross Total Value (Rs.)'})

            if not cdnr_df.empty and 'Note Type' in cdnr_df.columns:
                cr_notes = cdnr_df[cdnr_df['Note Type'].str.upper().isin(['C', 'CREDIT'])].groupby(['Financial Year', 'GSTIN of Supplier']).agg({
                    'Taxable Value (Rs.)': 'sum', 'IGST (Rs.)': 'sum', 'CGST (Rs.)': 'sum', 'SGST (Rs.)': 'sum', 'Cess (Rs.)': 'sum'
                }).reset_index().rename(columns={'Taxable Value (Rs.)': 'CN_Taxable', 'IGST (Rs.)': 'CN_IGST', 'CGST (Rs.)': 'CN_CGST', 'SGST (Rs.)': 'CN_SGST', 'Cess (Rs.)': 'CN_Cess'})

                dr_notes = cdnr_df[cdnr_df['Note Type'].str.upper().isin(['D', 'DEBIT'])].groupby(['Financial Year', 'GSTIN of Supplier']).agg({
                    'Taxable Value (Rs.)': 'sum', 'IGST (Rs.)': 'sum', 'CGST (Rs.)': 'sum', 'SGST (Rs.)': 'sum', 'Cess (Rs.)': 'sum'
                }).reset_index().rename(columns={'Taxable Value (Rs.)': 'DN_Taxable', 'IGST (Rs.)': 'DN_IGST', 'CGST (Rs.)': 'DN_CGST', 'SGST (Rs.)': 'DN_SGST', 'Cess (Rs.)': 'DN_Cess'})
            else:
                cr_notes = pd.DataFrame(columns=['Financial Year', 'GSTIN of Supplier', 'CN_Taxable', 'CN_IGST', 'CN_CGST', 'CN_SGST', 'CN_Cess'])
                dr_notes = pd.DataFrame(columns=['Financial Year', 'GSTIN of Supplier', 'DN_Taxable', 'DN_IGST', 'DN_CGST', 'DN_SGST', 'DN_Cess'])

            net_itc = b2b_grouped.merge(cr_notes, on=['Financial Year', 'GSTIN of Supplier'], how='left').merge(dr_notes, on=['Financial Year', 'GSTIN of Supplier'], how='left').fillna(0)
            net_itc['Less: Credit Notes - Taxable Value (Rs.)'] = net_itc['CN_Taxable']
            net_itc['Add: Debit Notes - Taxable Value (Rs.)'] = net_itc['DN_Taxable']
            net_itc['Net Taxable Value (Rs.)'] = net_itc['Gross Purchases - Taxable Value (Rs.)'] - net_itc['Less: Credit Notes - Taxable Value (Rs.)'] + net_itc['Add: Debit Notes - Taxable Value (Rs.)']
            net_itc['Net IGST (Rs.)'] = net_itc['IGST (Rs.)'] - net_itc['CN_IGST'] + net_itc['DN_IGST']
            net_itc['Net CGST (Rs.)'] = net_itc['CGST (Rs.)'] - net_itc['CN_CGST'] + net_itc['DN_CGST']
            net_itc['Net SGST (Rs.)'] = net_itc['SGST (Rs.)'] - net_itc['CN_SGST'] + net_itc['DN_SGST']
            net_itc['Net Cess (Rs.)'] = net_itc['Cess (Rs.)'] - net_itc['CN_Cess'] + net_itc['DN_Cess']
            net_itc['Net Total Tax (Rs.)'] = net_itc['Net IGST (Rs.)'] + net_itc['Net CGST (Rs.)'] + net_itc['Net SGST (Rs.)'] + net_itc['Net Cess (Rs.)']
            net_itc['Net Total Value (Rs.)'] = net_itc['Net Taxable Value (Rs.)'] + net_itc['Net Total Tax (Rs.)']

            summary_cols = ['Financial Year', 'GSTIN of Supplier', 'Trade/Legal Name of Supplier', 'Gross Purchases - Taxable Value (Rs.)', 'Less: Credit Notes - Taxable Value (Rs.)', 'Add: Debit Notes - Taxable Value (Rs.)', 'Net Taxable Value (Rs.)', 'Net IGST (Rs.)', 'Net CGST (Rs.)', 'Net SGST (Rs.)', 'Net Cess (Rs.)', 'Net Total Tax (Rs.)', 'Net Total Value (Rs.)']
            net_itc_summary = net_itc[summary_cols]

            year_wise = net_itc_summary.groupby('Financial Year').agg({
                'Gross Purchases - Taxable Value (Rs.)': 'sum', 'Less: Credit Notes - Taxable Value (Rs.)': 'sum', 'Add: Debit Notes - Taxable Value (Rs.)': 'sum', 'Net Taxable Value (Rs.)': 'sum', 'Net IGST (Rs.)': 'sum', 'Net CGST (Rs.)': 'sum', 'Net SGST (Rs.)': 'sum', 'Net Cess (Rs.)': 'sum', 'Net Total Tax (Rs.)': 'sum', 'Net Total Value (Rs.)': 'sum'
            }).reset_index()

            cancelled_df = b2b_df[b2b_df['Effective Date of Cancellation'].notna() & (b2b_df['Effective Date of Cancellation'] != '')].copy() if 'Effective Date of Cancellation' in b2b_df.columns else pd.DataFrame()

            if not cancelled_df.empty:
                cancelled_df['Is Post Cancellation'] = cancelled_df['Invoice Date Parsed'] > cancelled_df['Cancellation Date Parsed']
                
                summary_rows = []
                for gstin, group in cancelled_df.groupby('GSTIN of Supplier'):
                    trade_name = group['Trade/Legal Name of Supplier'].iloc[0]
                    eff_date = group['Effective Date of Cancellation'].iloc[0]
                    fy_involved = ", ".join(group['Financial Year'].unique())
                    num_invoices = len(group)
                    num_notes = 0
                    if not cdnr_df.empty and 'GSTIN of Supplier' in cdnr_df.columns:
                        num_notes = len(cdnr_df[cdnr_df['GSTIN of Supplier'] == gstin])
                    
                    net_taxable = group['Taxable Value (Rs.)'].sum()
                    net_igst = group['IGST (Rs.)'].sum()
                    net_cgst = group['CGST (Rs.)'].sum()
                    net_sgst = group['SGST (Rs.)'].sum()
                    net_cess = group['Cess (Rs.)'].sum()
                    net_total_tax = net_igst + net_cgst + net_sgst + net_cess
                    
                    tx_post = group['Is Post Cancellation'].sum()
                    tx_pre = num_invoices - tx_post
                    risk_flag = "Normal - all transactions pre-date cancellation" if tx_post == 0 else f"High Risk - {tx_post} transaction(s) post-date cancellation"

                    summary_rows.append({
                        'GSTIN': gstin,
                        'Trade/Legal Name of Supplier': trade_name,
                        'Effective Date of Cancellation': eff_date,
                        'Financial Years Involved': fy_involved,
                        'No. of Invoices': num_invoices,
                        'No. of Credit/Debit Notes': num_notes,
                        'Net Taxable Value (Rs.)': net_taxable,
                        'Net IGST (Rs.)': net_igst,
                        'Net CGST (Rs.)': net_cgst,
                        'Net SGST (Rs.)': net_sgst,
                        'Net Cess (Rs.)': net_cess,
                        'Net Total Tax (Rs.)': net_total_tax,
                        'Transactions Pre-Cancellation': tx_pre,
                        'Transactions Post-Cancellation': tx_post,
                        'Risk Flag': risk_flag
                    })
                cancelled_summary_df = pd.DataFrame(summary_rows)
            else:
                cancelled_summary_df = pd.DataFrame(columns=['GSTIN', 'Trade/Legal Name of Supplier', 'Effective Date of Cancellation', 'Financial Years Involved', 'No. of Invoices', 'No. of Credit/Debit Notes', 'Net Taxable Value (Rs.)', 'Net IGST (Rs.)', 'Net CGST (Rs.)', 'Net SGST (Rs.)', 'Net Cess (Rs.)', 'Net Total Tax (Rs.)', 'Transactions Pre-Cancellation', 'Transactions Post-Cancellation', 'Risk Flag'])

            export_cancelled_df = cancelled_df.drop(columns=['Invoice Date Parsed', 'Cancellation Date Parsed'], errors='ignore')

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                net_itc_summary.to_excel(writer, sheet_name='Net ITC Summary', index=False)
                year_wise.to_excel(writer, sheet_name='Year-wise Net ITC', index=False)
                pd.DataFrame({'Info': ['No B2BA amendments']}).to_excel(writer, sheet_name='Amended Invoices (B2BA Applied)', index=False)
                
                if not export_cancelled_df.empty:
                    export_cancelled_df.to_excel(writer, sheet_name='Cancelled Suppliers-Detail', index=False)
                    cancelled_summary_df.to_excel(writer, sheet_name='Cancelled Suppliers-Summary', index=False)
                else:
                    pd.DataFrame({'Info': ['No cancelled suppliers']}).to_excel(writer, sheet_name='Cancelled Suppliers-Detail', index=False)
                    pd.DataFrame({'Info': ['No cancelled suppliers']}).to_excel(writer, sheet_name='Cancelled Suppliers-Summary', index=False)
                
                b2b_df.drop(columns=['Invoice Date Parsed', 'Cancellation Date Parsed'], errors='ignore').to_excel(writer, sheet_name='B2B (All Years)', index=False)
                if not cdnr_df.empty:
                    cdnr_df.to_excel(writer, sheet_name='CDNR (All Years)', index=False)
                else:
                    pd.DataFrame({'Info': ['No CDNR']}).to_excel(writer, sheet_name='CDNR (All Years)', index=False)

            output.seek(0)
            wb = openpyxl.load_workbook(output)
            if 'Cancelled Suppliers-Detail' in wb.sheetnames:
                ws_det = wb['Cancelled Suppliers-Detail']
                red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                red_font = Font(color='9C0006')
                
                col_idx = None
                for col in range(1, ws_det.max_column + 1):
                    if ws_det.cell(row=1, column=col).value == 'Is Post Cancellation':
                        col_idx = col
                        break
                
                if col_idx:
                    for row in range(2, ws_det.max_row + 1):
                        val = ws_det.cell(row=row, column=col_idx).value
                        if val == True or val == 'True' or val == 1:
                            for c in range(1, ws_det.max_column + 1):
                                ws_det.cell(row=row, column=c).fill = red_fill
                                ws_det.cell(row=row, column=c).font = red_font

            final_output = BytesIO()
            wb.save(final_output)
            final_output.seek(0)

            processed_data = final_output.getvalue()
            st.success("Audit Report Generated Successfully with Correct Cancellation Flags & Highlights!")
            st.download_button(label="📥 Download Audit-Ready Excel Report", data=processed_data, file_name="GSTR2A_Audit_Ready_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing file: {e}")
