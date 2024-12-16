import os
import openpyxl
import xlrd
from openpyxl.utils import get_column_letter

def split_excel_by_elements(input_file_path, output_directory_path):
    """
    Excelファイルに含まれるすべてのシートを個別のファイルに分割する関数。
    列幅調整機能を追加。
    """
    try:
        if input_file_path.endswith('.xls'):
            workbook = xlrd.open_workbook(input_file_path)
            is_xls = True
        elif input_file_path.endswith('.xlsx'):
            workbook = openpyxl.load_workbook(input_file_path)
            is_xls = False
        else:
            print(f"Unsupported file format: {input_file_path}")
            return

        if is_xls:
            for sheet in workbook.sheets():
                new_workbook = openpyxl.Workbook()
                new_sheet = new_workbook.active
                for rx in range(sheet.nrows):
                    new_sheet.append(sheet.row_values(rx))

                # 列幅を調整
                for col_idx in range(sheet.ncols):
                    column_letter = get_column_letter(col_idx + 1)
                    new_sheet.column_dimensions[column_letter].width = 20  # 例：幅20に設定
                    # もしくは、最大文字数に合わせて調整する場合：
                    # max_length = 0
                    # for cell in sheet.col(col_idx):
                    #     try:
                    #         if len(str(cell.value)) > max_length:
                    #             max_length = len(str(cell.value))
                    #     except:  # cell.valueがNoneの場合などをキャッチ
                    #         pass
                    # new_sheet.column_dimensions[column_letter].width = max_length + 2

                output_file_path = os.path.join(output_directory_path, f"{sheet.name}.xlsx")
                new_workbook.save(output_file_path)
        else:
            for sheet in workbook.worksheets:
                new_workbook = openpyxl.Workbook()
                new_sheet = new_workbook.active
                for row in sheet.iter_rows():
                    new_sheet.append([cell.value for cell in row])

                # 列幅を調整
                for col in new_sheet.columns:
                    max_length = 0
                    column = list(col)
                    try:
                        for cell in column:
                            if cell.value:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                    except AttributeError: # セルに値がない場合のエラーをキャッチ
                        pass
                    adjusted_width = (max_length + 2)
                    if adjusted_width < 10: # 最小幅を設定
                         adjusted_width = 10
                    new_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

                output_file_path = os.path.join(output_directory_path, f"{sheet.title}.xlsx")
                new_workbook.save(output_file_path)

    except FileNotFoundError:
        print(f"Input file not found: {input_file_path}")
    except Exception as e:
        print(f"An error occurred processing {input_file_path}: {e}")

def process_directory(input_directory_path, output_directory_path):
    """
    指定されたディレクトリ内のExcelファイルを処理する関数。

    Args:
        input_directory_path: 入力ディレクトリのパス
        output_directory_path: 出力ディレクトリのパス
    """
    if not os.path.exists(output_directory_path):
        os.makedirs(output_directory_path)
    for filename in os.listdir(input_directory_path):
        if filename.endswith(('.xlsx', '.xls')):
            file_path = os.path.join(input_directory_path, filename)
            split_excel_by_elements(file_path, output_directory_path)

# 使用例
input_directory = "katashiki"
output_directory = "output"
process_directory(input_directory, output_directory)
