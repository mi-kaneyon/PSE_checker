from PySide2.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QWidget, QLabel, QLineEdit, QHBoxLayout, QPushButton, QVBoxLayout
from PySide2.QtGui import QColor
import sys
import pandas as pd

class DenanTool(QMainWindow):
    def __init__(self):
        super(DenanTool, self).__init__()
        self.initUI()
        self.specified_products = pd.DataFrame()
        self.non_specified_products = pd.DataFrame()

    def initUI(self):
        # Window setup
        self.setWindowTitle('電気用品安全法 検索ツール')
        self.setGeometry(100, 100, 800, 600)
        
        # Table setup
        self.table = QTableWidget()
        self.table.setColumnCount(3)  # Number of columns
        self.table.setHorizontalHeaderLabels(['Product', 'explain', 'PSELOGO'])
        
        # Search bar and buttons setup
        self.search_bar = QLineEdit(self)
        self.search_button = QPushButton('Search', self)
        self.manual_search_button = QPushButton('Manual Search', self)
        self.search_button.clicked.connect(self.search_products)
        self.manual_search_button.clicked.connect(self.manual_search)
        search_layout = QHBoxLayout()
        search_layout.addWidget(self.search_bar)
        search_layout.addWidget(self.search_button)
        search_layout.addWidget(self.manual_search_button)
        
        # Main layout setup
        main_layout = QVBoxLayout()
        main_layout.addLayout(search_layout)
        main_layout.addWidget(self.table)
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)
        
        # Adjust column widths after UI setup
        self.adjust_column_widths()

    def adjust_column_widths(self):
        # Set 'explain' column width to 35% of table width
        self.table.setColumnWidth(1, int(self.table.width() * 0.35))
        
    def load_data_from_excel(self):
        self.specified_products = pd.read_excel('specified_electrical_products.xlsx')
        self.non_specified_products = pd.read_excel('non-specified_electrical_products.xlsx')
        self.populate_table()

    def populate_table(self):
        self.table.setRowCount(0)
        combined_data = pd.concat([self.specified_products, self.non_specified_products])
        for index, row in combined_data.iterrows():
            row_pos = self.table.rowCount()
            self.table.insertRow(row_pos)
            self.table.setItem(row_pos, 0, QTableWidgetItem(str(row['Product'])))
            self.table.setItem(row_pos, 1, QTableWidgetItem(str(row['explain'])))
            pse_item = QTableWidgetItem(row['PSELOGO'])
            pse_item.setBackground(QColor(255, 0, 0) if row['PSELOGO'] == '<PSE>' else QColor(0, 0, 255))
            pse_item.setForeground(QColor(255, 255, 255))
            self.table.setItem(row_pos, 2, pse_item)
        self.table.resizeColumnsToContents()

    def search_products(self):
        search_text = self.search_bar.text().lower()
        # 3文字以上であることを確認


        # テーブルの全ての行をループ
        for i in range(self.table.rowCount()):
            row_contains_search_text = False
            # 各行の全てのセルをループ
            for j in range(self.table.columnCount()):
                cell_text = self.table.item(i, j).text().lower() if self.table.item(i, j) else ''
                # セルの文字列に検索テキストが部分的に含まれているかを確認
                if search_text in cell_text or cell_text in search_text:
                    row_contains_search_text = True
                    break
            # 検索テキストが含まれているかに応じて行の表示・非表示を切り替え
            self.table.setRowHidden(i, not row_contains_search_text)

                
    def manual_search(self):
        search_text = self.search_bar.text() + ' PSE マーク 対象'
        print(search_text)  # Placeholder for manual search functionality
        
    def update_pse_mark_label(self, row, column):
        if column == 2:
            pse_mark = self.table.item(row, column).text()
            self.pse_mark_label.setText('PSE Mark: ' + pse_mark)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWin = DenanTool()
    mainWin.load_data_from_excel()
    mainWin.show()
    sys.exit(app.exec_())
