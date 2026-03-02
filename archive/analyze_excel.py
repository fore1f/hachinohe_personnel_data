import pandas as pd
import sys

def main():
    file_path = "C:\\Users\\tked1\\py\\hachinohe\\①R7.4.1八戸市　人事異動内示データ　.xlsx"
    try:
        df = pd.read_excel(file_path, header=None)
        
        with open("analysis_output.txt", "w", encoding="utf-8") as f:
            f.write("--- Excelデータの構成（先頭50行） ---\n")
            for index, row in df.head(50).iterrows():
                row_list = row.tolist()
                clean_row = [str(cell).replace('\n', ' ').replace('\r', '') if pd.notna(cell) else '' for cell in row_list]
                f.write(f"Row {index}: {clean_row}\n")
            
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
