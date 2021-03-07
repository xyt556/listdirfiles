import sys
from PyQt5.QtWidgets import QApplication, QWidget, QAction, QMainWindow, QLineEdit, QGridLayout, QLabel, QPushButton, QFileDialog, QDesktopWidget
from PyQt5.QtGui import QIcon, QDesktopServices
from PyQt5.QtCore import QUrl
from PyQt5 import QtWidgets
import os
import listdirfiles as ldf


class ListFilesGui(QWidget):

    def __init__(self):
        super().__init__()
        self.grid = QGridLayout()
        self.result_file = QLabel()
        self.generate_btn = QPushButton('生成')
        self.save_filename_edit = QLineEdit()
        self.file_suffix_edit = QLineEdit()
        self.root_path_edit = QLineEdit()
        self.root_path_lbl = QLabel('目录：')
        self.file_suffix_lbl = QLabel('文件后缀：')
        self.save_filename_lbl = QLabel('保存文件名：')
        self.result_file_lbl = QLabel('结果文件：')
        self.open_btn = QPushButton('打开')
        self.file_path = ''
        self.init_ui()

    def init_ui(self):
        self.resize(600, 200)
        self.setWindowTitle('目录列表生成工具')
        self.setWindowIcon(QIcon('dir.png'))

        # 定义标签名称

        # 定义输入框
        self.root_path_edit.setPlaceholderText('请输入需要生成文件列表文件的根目录')
        # self.root_path_btn = QPushButton('浏览')
        # self.root_path_btn.setCheckable(False)

        self.file_suffix_edit.setPlaceholderText('请输入文件后缀名(不带.)，多个类型请用英文逗号分隔，不区分后缀输入all')

        self.save_filename_edit.setPlaceholderText('请输入生成的excel文件名称，默认default.xlsx')

        # 信号绑定
        self.generate_btn.clicked.connect(self.generate_file)

        self.open_btn.clicked.connect(self.open_res_file)

        # grid布局，将控件塞进去
        self.setLayout(self.grid)

        self.grid.addWidget(self.root_path_lbl, 1, 0)
        self.grid.addWidget(self.root_path_edit, 1, 1)

        self.grid.addWidget(self.file_suffix_lbl, 2, 0)
        self.grid.addWidget(self.file_suffix_edit, 2, 1)

        self.grid.addWidget(self.save_filename_lbl, 3, 0)
        self.grid.addWidget(self.save_filename_edit, 3, 1)

        self.grid.addWidget(self.generate_btn, 4, 0, 1, 0)

        self.grid.addWidget(self.result_file_lbl, 5, 0)
        self.grid.addWidget(self.result_file, 5, 1)

        self.grid.addWidget(self.open_btn, 6, 0, 1, 0)


        self.show()

    def generate_file(self):
        root_path = self.root_path_edit.text()
        suffix = self.file_suffix_edit.text()
        filename = self.save_filename_edit.text()

        content = ldf.get_all_files(path=root_path, suffix=suffix)
        self.file_path = ldf.write_to_excel(content=content, filename=filename, path=root_path)
        # QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        self.result_file.setText(self.file_path)

    def open_res_file(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.file_path))


if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = ListFilesGui()
    sys.exit(app.exec_())
