import os
import xlrd
import openpyxl

def list_katashiki_sheets(input_directory_path):
    """
    katashikiディレクトリ内のすべてのシート名を取得（インデックス除外）。
    """
    katashiki_sheets = []

    for filename in os.listdir(input_directory_path):
        file_path = os.path.join(input_directory_path, filename)
        if filename.endswith('.xls'):
            workbook = xlrd.open_workbook(file_path)
            valid_sheets = [sheet for sheet in workbook.sheet_names() if "インデックス" not in sheet]
            katashiki_sheets.extend(valid_sheets)
        elif filename.endswith('.xlsx'):
            workbook = openpyxl.load_workbook(file_path, read_only=True)
            valid_sheets = [sheet for sheet in workbook.sheetnames if "インデックス" not in sheet]
            katashiki_sheets.extend(valid_sheets)

    return katashiki_sheets

def list_output_files(output_directory_path):
    """
    outputディレクトリ内のファイル名（拡張子なし）を取得。
    """
    return [os.path.splitext(file)[0] for file in os.listdir(output_directory_path) if file.endswith('.xlsx')]

def find_missing_files(katashiki_sheets, output_files):
    """
    katashikiのシート名とoutputファイル名を比較し、不足分を特定。
    """
    missing_files = set(katashiki_sheets) - set(output_files)
    return missing_files

# 実行
katashiki_directory = "katashiki"
output_directory = "output"

katashiki_sheets = list_katashiki_sheets(katashiki_directory)
output_files = list_output_files(output_directory)

# 不足ファイルを特定
missing_files = find_missing_files(katashiki_sheets, output_files)

print(f"Total missing files: {len(missing_files)}")
print("Missing file names:")
for missing in missing_files:
    print(missing)
