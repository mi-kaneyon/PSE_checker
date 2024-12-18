import os
import xlrd
import openpyxl

def count_sheets_in_directory(directory_path):
    """
    katashikiディレクトリ内のExcelファイルからインデックスシートを除いたシート数をカウントする。
    """
    total_files = 0
    total_sheets = 0

    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            total_files += 1
            print(f"Processing file: {filename}")
            try:
                if filename.endswith('.xls'):
                    workbook = xlrd.open_workbook(file_path)
                    valid_sheets = [sheet for sheet in workbook.sheet_names() if "インデックス" not in sheet]
                    total_sheets += len(valid_sheets)
                else:
                    workbook = openpyxl.load_workbook(file_path, read_only=True)
                    valid_sheets = [sheet for sheet in workbook.sheetnames if "インデックス" not in sheet]
                    total_sheets += len(valid_sheets)
            except Exception as e:
                print(f"Error processing file {filename}: {e}")

    print(f"\nSummary:")
    print(f"Total files processed: {total_files}")
    print(f"Total valid sheets (excluding インデックス): {total_sheets}")

# 実行
input_directory = "katashiki"
count_sheets_in_directory(input_directory)
