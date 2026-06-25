import os
import glob
import pandas as pd

file_list = glob.glob("*.xls")

for file_path in file_list:
    try:
        df_raw = pd.read_excel(file_path, header=None)
        
        header_idx = None
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.contains("No. Faktur").any():
                header_idx = idx
                break
        
        if header_idx is None:
            print(f"--> Header 'No. Faktur' tidak ditemukan di {file_path}")
            continue
            
        columns = [str(c).replace('\xa0', ' ').strip() for c in df_raw.iloc[header_idx]]
        df_raw.columns = columns
        
        df = df_raw.iloc[header_idx + 1:].copy()
        
        cols = [
            "No. Faktur", "Tgl Faktur", "Umur", "Kode", 
            "Nama Pelanggan", "Nilai Faktur", "Terutang", 
            "Nama Gudang", "Nama Penjual"
        ]
        
        df = df.dropna(subset=cols)
        df = df[df["No. Faktur"].astype(str).str.strip() != "No. Faktur"]
        
        df["No. Faktur"] = df["No. Faktur"].astype(str).apply(lambda x: x.split(",")[0].replace(".", "").strip())
        df["No. Faktur"] = pd.to_numeric(df["No. Faktur"], errors='coerce').fillna(0).astype(int)
        
        df["Umur"] = df["Umur"].astype(str).apply(lambda x: x.split(",")[0].replace(".", "").strip())
        df["Umur"] = pd.to_numeric(df["Umur"], errors='coerce').fillna(0).astype(int)
        
        for col in ["Nilai Faktur", "Terutang"]:
            df[col] = df[col].astype(str).apply(lambda x: x.split(",")[0].replace(".", "").strip())
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        
        df_clean = df[cols].copy()
        
        grouped = (
            df_clean.groupby(["Kode", "Nama Pelanggan"])
            .agg({"Umur": "mean", "Nilai Faktur": "mean", "No. Faktur": "count"})
            .reset_index()
        )
        
        grouped.columns = [
            "No. Pelanggan",
            "Nama Pelanggan",
            "AVG UMUR PIUTANG",
            "AVG NILAI FAKTUR",
            "JUMLAH INVOICE",
        ]
        grouped = grouped.sort_values(by="No. Pelanggan")
        
        grouped["AVG UMUR PIUTANG"] = pd.to_numeric(grouped["AVG UMUR PIUTANG"], errors='coerce').fillna(0).round().astype(int)
        grouped["AVG NILAI FAKTUR"] = pd.to_numeric(grouped["AVG NILAI FAKTUR"], errors='coerce').fillna(0).round().astype(int)
        grouped["JUMLAH INVOICE"] = grouped["JUMLAH INVOICE"].astype(int)
        
        base_name = os.path.splitext(file_path)[0]
        output_file = f"{base_name}_hasil.xlsx"
        
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df_clean.to_excel(writer, index=False, sheet_name="Data Bersih")
            grouped.to_excel(writer, index=False, sheet_name="Ringkasan")
            
            worksheet_clean = writer.sheets["Data Bersih"]
            for col_num, col_name in enumerate(df_clean.columns, 1):
                if col_name in ["Nilai Faktur", "Terutang"]:
                    for row in range(2, worksheet_clean.max_row + 1):
                        worksheet_clean.cell(row=row, column=col_num).number_format = '#,##0'
            
            for col in worksheet_clean.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                worksheet_clean.column_dimensions[col_letter].width = max(max_len + 3, 10)
                
            worksheet_ringkasan = writer.sheets["Ringkasan"]
            for col_num, col_name in enumerate(grouped.columns, 1):
                if col_name in ["AVG NILAI FAKTUR"]:
                    for row in range(2, worksheet_ringkasan.max_row + 1):
                        worksheet_ringkasan.cell(row=row, column=col_num).number_format = '#,##0'
                        
            for col in worksheet_ringkasan.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                worksheet_ringkasan.column_dimensions[col_letter].width = max(max_len + 3, 10)
        
        print(f"--> Berhasil memproses: {file_path} menjadi {output_file}")
    except Exception as e:
        print(f"--> Gagal memproses {file_path}: {str(e)}")