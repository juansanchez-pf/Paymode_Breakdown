import pandas as pd
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import time
import zipfile
from google.colab import files # Importante: Librerías específicas de Colab

def format_excel_file(filename, payout_df, processed_df):
    """Applies detailed formatting to the generated Excel file."""
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        if not payout_df.empty:
            payout_df.to_excel(writer, sheet_name='Payout', index=False)
        if not processed_df.empty:
            processed_df.to_excel(writer, sheet_name='Processed', index=False)
    
    workbook = openpyxl.load_workbook(filename)

    header_font = Font(bold=True)
    total_row_font = Font(bold=True)
    header_fill = PatternFill(start_color="C5D9F1", end_color="C5D9F1", fill_type="solid")
    thin_side = Side(style='thin')
    accounting_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        if sheet.max_row == 0: continue
        
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill

        for col in sheet.columns:
            max_length = 0
            column_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                except: pass
            sheet.column_dimensions[column_letter].width = (max_length + 2)

        if sheet.max_row > 1:
            for cell in sheet[sheet.max_row]:
                cell.font = total_row_font
                if cell.value: cell.fill = header_fill
        
        max_col, max_row = sheet.max_column, sheet.max_row

        if max_row >= 1:
            for col_idx in range(1, max_col + 1):
                cell = sheet.cell(row=1, column=col_idx)
                border = Border(top=thin_side, bottom=thin_side)
                if col_idx == 1: border.left = thin_side
                if col_idx == max_col: border.right = thin_side
                cell.border = border
        
        if max_row > 2:
            data_min_row, data_max_row = 2, max_row - 1
            for row in sheet.iter_rows(min_row=data_min_row, max_row=data_max_row, min_col=1, max_col=max_col):
                for cell in row:
                    border = Border()
                    if cell.row == data_min_row: border.top = thin_side
                    if cell.row == data_max_row: border.bottom = thin_side
                    if cell.column == 1: border.left = thin_side
                    if cell.column == max_col: border.right = thin_side
                    cell.border = border
        
        if max_row > 1:
            for cell in sheet[max_row]:
                if cell.value: cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        headers = [cell.value for cell in sheet[1]]
        cols_to_format = ["Payment Amount", "Coupa Customer Dividend"] if sheet_name == 'Payout' else ["Payment Amount"]
        for col_name in cols_to_format:
            if col_name in headers:
                col_idx = headers.index(col_name) + 1
                for cell in sheet[get_column_letter(col_idx)]:
                    if cell.row > 1: cell.number_format = accounting_format
    workbook.save(filename)

def main():
    """Main function adapted for Google Colab."""
    start_time = time.time()

    # --- 1. Pedir al usuario que suba el archivo ---
    print("Por favor, sube el archivo 'Coupa Paymode-X Dividends Report.xlsx'")
    uploaded = files.upload()
    
    input_filename = next(iter(uploaded))
    print(f"Archivo '{input_filename}' subido exitosamente.")

    # --- 2. Crear una carpeta temporal para los resultados ---
    output_folder = "Px_Customers_Breakdown"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        xls = pd.ExcelFile(input_filename)
        payout_df_full = pd.read_excel(xls, sheet_name="Payout")
        processed_df_full = pd.read_excel(xls, sheet_name="Processed")
    except Exception as e:
        print(f"Ocurrió un error leyendo el archivo Excel: {e}")
        return

    payout_df_full.columns = payout_df_full.columns.str.strip()
    processed_df_full.columns = processed_df_full.columns.str.strip()

    rename_dict = {
        'Disburser Paymode-X Account': 'Disburser Paymode Account',
        'Collector Paymode-X Account': 'Collector Paymode'
    }
    payout_df_full.rename(columns=rename_dict, inplace=True)
    processed_df_full.rename(columns=rename_dict, inplace=True)

    payout_cols_desired = [
        "Disburser Company Name", "Disburser Paymode Account", "Collector Paymode", 
        "Collector Network Fee Billing Method", "Channel Dividend Currency", "DPA", 
        "Payment Credit Settlement Date", "Date Fees Collected", "Payment Amount", 
        "Coupa Customer Dividend"
    ]
    processed_cols_desired = [
        "Disburser Company Name", "Disburser Paymode Account", "Collector Paymode", 
        "Collector Network Fee Billing Method", "Channel Dividend", "Currency", "DPA", 
        "Payment Credit Settlement Date", "Date Fees Collected", "Payment Amount", "Fee Details"
    ]
    
    # El resto del script de procesamiento es idéntico
    actual_payout_cols = payout_df_full.columns.tolist()
    payout_cols_to_keep = [col for col in payout_cols_desired if col in actual_payout_cols]
    payout_df_full = payout_df_full[payout_cols_to_keep]

    actual_processed_cols = processed_df_full.columns.tolist()
    processed_cols_to_keep = [col for col in processed_cols_desired if col in actual_processed_cols]
    processed_df_full = processed_df_full[processed_cols_to_keep]

    all_customers = pd.concat([
        payout_df_full["Disburser Company Name"],
        processed_df_full["Disburser Company Name"]
    ]).dropna().unique()

    generated_files = []
    for customer in all_customers:
        print(f"Procesando cliente: {customer}")
        customer_payout_df = payout_df_full[payout_df_full["Disburser Company Name"] == customer].copy()
        customer_processed_df = processed_df_full[processed_df_full["Disburser Company Name"] == customer].copy()

        if not customer_payout_df.empty:
            payout_totals = {"Disburser Company Name": "Total"}
            if "Payment Amount" in customer_payout_df.columns:
                payout_totals["Payment Amount"] = customer_payout_df["Payment Amount"].sum()
            if "Coupa Customer Dividend" in customer_payout_df.columns:
                payout_totals["Coupa Customer Dividend"] = customer_payout_df["Coupa Customer Dividend"].sum()
            payout_total_row = pd.DataFrame([payout_totals])
            customer_payout_df = pd.concat([customer_payout_df, payout_total_row], ignore_index=True)

        if not customer_processed_df.empty:
            processed_totals = {"Disburser Company Name": "Total"}
            if "Payment Amount" in customer_processed_df.columns:
                processed_totals["Payment Amount"] = customer_processed_df["Payment Amount"].sum()
            processed_total_row = pd.DataFrame([processed_totals])
            customer_processed_df = pd.concat([customer_processed_df, processed_total_row], ignore_index=True)
            
        safe_customer_name = "".join(c for c in customer if c.isalnum() or c in (' ', '.', '_')).rstrip()
        output_filename = os.path.join(output_folder, f"{safe_customer_name} Monthly Dividend Report.xlsx")
        
        format_excel_file(output_filename, customer_payout_df, customer_processed_df)
        generated_files.append(output_filename)

    # --- 3. Comprimir los resultados y ofrecer la descarga ---
    zip_filename = "Reportes_Por_Cliente.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for file in generated_files:
            zipf.write(file)
    
    print(f"\nProceso completado. Descargando '{zip_filename}' con todos los reportes.")
    files.download(zip_filename)

    end_time = time.time()
    total_time = end_time - start_time
    minutes, seconds = int(total_time // 60), int(total_time % 60)
    print(f"Tiempo total de ejecución: {minutes} minutos y {seconds} segundos.")

if __name__ == "__main__":
    main()
