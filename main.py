# -*- coding:utf-8 -*-
import json
import os
import shutil
import sys
import time
import base64
import re
from datetime import datetime
import sqlite3
import queue
import threading
import xlrd
import xlwt
import pandas as pd

from DrissionPage import ChromiumPage, ChromiumOptions
ChromiumOptions().set_browser_path("C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe").save()

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import QtCore, QtGui
from PyQt5.QtGui import QStandardItem, QIcon, QIntValidator
from qt_material import apply_stylesheet


class FileHandle:
    def __init__(self, frame):
        self.frame = frame

    def factory_excel_split(self, file_path, save_path):
        """不同工厂的信息分割成不同的excel"""
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)

        all_info_dict = {}
        start_row = 11
        for row in range(sheet.nrows):
            if row < start_row:
                continue
            item = sheet.cell_value(row, 1)
            if item == "":
                continue
            factory_id = item.split("-")[0]
            if factory_id in all_info_dict.keys():
                data_list = all_info_dict[factory_id]
            else:
                data_list = []
                all_info_dict[factory_id] = data_list
            info = {}
            info["item"] = item
            info["price_usd"] = sheet.cell_value(row, 3)
            info["item_size_l"] = sheet.cell_value(row, 4)
            info["item_size_w"] = sheet.cell_value(row, 5)
            info["item_size_h"] = sheet.cell_value(row, 6)
            info["inner_pack"] = sheet.cell_value(row, 7)
            info["master_pack"] = sheet.cell_value(row, 8)
            info["carton_cbm"] = sheet.cell_value(row, 9)
            info["carton_l"] = sheet.cell_value(row, 10)
            info["carton_w"] = sheet.cell_value(row, 11)
            info["carton_h"] = sheet.cell_value(row, 12)
            info["n_w_kgs"] = sheet.cell_value(row, 13)
            info["g_w_kgs"] = sheet.cell_value(row, 14)
            info["price_rmb"] = sheet.cell_value(row, 15)
            data_list.append(info)
        print(111)
        self.frame.signal_log.emit("要处理的文件已经读取好了，现在开始写入新文件，再等一下哦")
        template_path = os.path.join("Resource", "template.xls")
        tempplate_workbook = xlrd.open_workbook(template_path)
        template_sheet = tempplate_workbook.sheet_by_index(0)
        for factory_id, data_list in all_info_dict.items():
            new_workbook = xlwt.Workbook()
            new_sheet = new_workbook.add_sheet('Sheet1')
            # 复制数据
            for row in range(template_sheet.nrows):
                for col in range(template_sheet.ncols):
                    value = template_sheet.cell_value(row, col)
                    new_sheet.write(row, col, value)

            # 写入新数据
            current_row = 11
            for info in data_list:
                new_sheet.write(current_row, 1, info["item"])
                new_sheet.write(current_row, 3, info["price_usd"])
                new_sheet.write(current_row, 4, info["item_size_l"])
                new_sheet.write(current_row, 5, info["item_size_w"])
                new_sheet.write(current_row, 6, info["item_size_h"])
                new_sheet.write(current_row, 7, info["inner_pack"])
                new_sheet.write(current_row, 8, info["master_pack"])
                new_sheet.write(current_row, 9, info["carton_cbm"])
                new_sheet.write(current_row, 10, info["carton_l"])
                new_sheet.write(current_row, 11, info["carton_w"])
                new_sheet.write(current_row, 12, info["carton_h"])
                new_sheet.write(current_row, 13, info["n_w_kgs"])
                new_sheet.write(current_row, 14, info["g_w_kgs"])
                new_sheet.write(current_row, 15, info["price_rmb"])
                current_row += 1
            file_name = str(factory_id) + ".xls"
            file_path = os.path.join(save_path, file_name)
            new_workbook.save(file_path)
            self.frame.signal_log.emit("%s已经保存好啦！"%file_name)
        self.frame.signal_log.emit("搞定^_^")




class ConfigSave:
    """配置保存类"""
    def __init__(self, frame):
        self.frame = frame
        self.sql = None
        self.db_name = "Config.db"
        self.table_name = "config"
        self.cursor = None
        self.init_table()

    def connect(self):
        self.sql = sqlite3.connect('Config.db')
        # 创建一个游标对象来执行SQL语句
        self.cursor = self.sql.cursor()

    def init_table(self):
        self.connect()
        self.cursor.execute('''SELECT count(name) FROM sqlite_master WHERE type='table' AND name='%s' ''' % self.table_name)
        # 如果表不存在，则创建表
        if self.cursor.fetchone()[0] == 0:
            self.cursor.execute('''CREATE TABLE "%s" (Name TEXT PRIMARY KEY, Value TEXT)''' % self.table_name)
            # 提交更改
            self.sql.commit()
        self.sql.close()

    def delete_table(self):
        self.connect()
        self.cursor.execute("DROP TABLE IF EXISTS " + self.table_name)
        self.sql.commit()
        self.sql.close()

    def clear_table(self):
        self.connect()
        self.cursor.execute("DELETE FROM " + self.table_name)
        self.sql.commit()
        self.sql.close()

    def save_config(self, name_list):
        """保存所有配置"""
        for name in name_list:
            self.save_single_config(name)

    def load_config(self, name_list):
        for name in name_list:
            self.load_single_config(name)

    def save_single_config(self, name):
        """保存单个配置"""
        # 值预处理
        control = getattr(self.frame, name)
        if  isinstance(control, QLineEdit):
            value = control.text()
        elif isinstance(control, QComboBox):
            value = control.currentText()
        else:
            value = control.text()
        content = json.dumps(value)

        # 保存到数据库
        self.connect()
        self.cursor.execute("SELECT Value FROM %s WHERE Name=?"%self.table_name, (name,))
        result = self.cursor.fetchone()
        if result:
            self.cursor.execute("UPDATE %s SET Value=? WHERE Name=?"%self.table_name, (content, name))
        else:
            self.cursor.execute("INSERT INTO %s (Name, Value) VALUES (?, ?)"%self.table_name,(name, content))
        self.sql.commit()
        self.sql.close()

    def load_single_config(self, name):
        """加载配置"""
        # 从数据库查询
        self.connect()
        self.cursor.execute("SELECT Value FROM %s WHERE Name=?" % self.table_name, (name,))
        result = self.cursor.fetchone()
        if result:
            content = result[0]
            value = json.loads(content)
        else:
            value = None

        # 赋值到控件
        if value is None:
            return
        control = getattr(self.frame, name)
        if isinstance(control, QLineEdit):
            control.setText(value)
        elif isinstance(control, QComboBox):
            control.setCurrentText(value)
        else:
            control.setText(value)


class MainFrame(QTabWidget):
    """界面类"""

    signal_hint_error = QtCore.pyqtSignal(str)
    signal_hint_info = QtCore.pyqtSignal(str)
    signal_log = QtCore.pyqtSignal(str)

    def __init__(self):
        super(MainFrame, self).__init__()
        self.file_handle = FileHandle(self)
        self.config_save = ConfigSave(self)
        self.website_entry_info = {}
        self.init_ui()
        self.log("日志系统")
        self.init_save()

        # 信号绑定
        self.signal_log.connect(self.log)
        self.signal_hint_info.connect(self.hint_infomation)
        self.signal_hint_error.connect(self.hint_error)


    def init_ui(self):
        self.setWindowTitle(" ")
        self.setWindowIcon(QtGui.QIcon("ico.ico"))
        self.resize(600, 400)

        # 拆分功能
        self.widget1 = QWidget(self)

        vbox = QVBoxLayout(self.widget1)
        self.widget1.setLayout(vbox)

        hbox_start = QHBoxLayout(self.widget1)
        self.start_button = QPushButton("开始处理", self.widget1)
        self.start_button.clicked.connect(self.start_handle)
        hbox_start.addWidget(self.start_button)
        vbox.addLayout(hbox_start)

        hbox_path = QHBoxLayout(self.widget1)
        self.file_path_ctrl = QLineEdit(self.widget1)
        self.file_path_ctrl.setEnabled(False)
        self.file_path_ctrl.setStyleSheet("color: #3CB371;")
        self.file_path_ctrl.setPlaceholderText("选择要处理的文件")
        self.file_path_choice_button = QPushButton("选择文件", self.widget1)
        self.file_path_choice_button.clicked.connect(self.choice_file_path)
        hbox_path.addWidget(self.file_path_ctrl)
        hbox_path.addWidget(self.file_path_choice_button)
        vbox.addLayout(hbox_path)

        hbox_path = QHBoxLayout(self.widget1)
        self.save_path_ctrl = QLineEdit(self.widget1)
        self.save_path_ctrl.setEnabled(False)
        self.save_path_ctrl.setStyleSheet("color: #3CB371;")
        self.save_path_ctrl.setPlaceholderText("文件存放路径")
        self.save_path_choice_button = QPushButton("选择路径", self.widget1)
        self.save_path_choice_button.clicked.connect(self.choice_save_path)
        hbox_path.addWidget(self.save_path_ctrl)
        hbox_path.addWidget(self.save_path_choice_button)
        vbox.addLayout(hbox_path)

        hbox_log = QHBoxLayout(self.widget1)
        self.log_ctrl = QTextEdit(self.widget1)
        self.log_ctrl.setReadOnly(True)
        self.log_ctrl.setStyleSheet("color: #3CB371;")
        hbox_log.addWidget(self.log_ctrl)
        vbox.addLayout(hbox_log)

        self.addTab(self.widget1, "拆分")


        # 提取货号
        self.widget2 = QWidget(self)
        vbox2 = QVBoxLayout(self.widget2)
        self.widget1.setLayout(vbox2)

        hbox_start = QHBoxLayout(self.widget2)
        self.start_button2 = QPushButton("开始处理", self.widget2)
        self.start_button2.clicked.connect(self.start_handle2)
        hbox_start.addWidget(self.start_button2)
        vbox2.addLayout(hbox_start)

        hbox_path = QHBoxLayout(self.widget2)
        self.file_path_ctrl2 = QLineEdit(self.widget2)
        self.file_path_ctrl2.setEnabled(False)
        self.file_path_ctrl2.setStyleSheet("color: #3CB371;")
        self.file_path_ctrl2.setPlaceholderText("选择要处理的文件")
        self.file_path_choice_button2 = QPushButton("选择文件", self.widget2)
        self.file_path_choice_button2.clicked.connect(self.choice_file_path2)
        hbox_path.addWidget(self.file_path_ctrl2)
        hbox_path.addWidget(self.file_path_choice_button2)
        vbox2.addLayout(hbox_path)

        hbox_path = QHBoxLayout(self.widget2)
        self.save_path_ctrl2 = QLineEdit(self.widget2)
        self.save_path_ctrl2.setEnabled(False)
        self.save_path_ctrl2.setStyleSheet("color: #3CB371;")
        self.save_path_ctrl2.setPlaceholderText("文件存放路径")
        self.save_path_choice_button2 = QPushButton("选择路径", self.widget2)
        self.save_path_choice_button2.clicked.connect(self.choice_save_path2)
        hbox_path.addWidget(self.save_path_ctrl2)
        hbox_path.addWidget(self.save_path_choice_button2)
        vbox2.addLayout(hbox_path)

        hbox_log = QHBoxLayout(self.widget2)
        self.log_ctrl2 = QTextEdit(self.widget2)
        self.log_ctrl2.setReadOnly(True)
        self.log_ctrl2.setStyleSheet("color: #3CB371;")
        hbox_log.addWidget(self.log_ctrl2)
        vbox2.addLayout(hbox_log)

        self.addTab(self.widget2, "提取货号")

        # 网站录入
        self.widget3 = QWidget(self)
        vbox2 = QVBoxLayout(self.widget3)
        self.widget1.setLayout(vbox2)

        hbox_start = QHBoxLayout(self.widget3)

        self.open_website_button = QPushButton("打开网站", self.widget3)
        self.open_website_button.clicked.connect(self.open_website)
        hbox_start.addWidget(self.open_website_button)

        self.current_row_ctrl = QLineEdit(self.widget3)
        self.current_row_ctrl.setText("3")
        self.current_row_ctrl.setFixedWidth(50)
        self.current_row_ctrl.setStyleSheet("color: #3CB371;")
        hbox_start.addWidget(self.current_row_ctrl)

        self.entry_button = QPushButton("录入", self.widget3)
        self.entry_button.clicked.connect(self.website_entry)
        hbox_start.addWidget(self.entry_button)
        vbox2.addLayout(hbox_start)

        hbox_path = QHBoxLayout(self.widget3)
        self.file_path_ctrl3 = QLineEdit(self.widget3)
        self.file_path_ctrl3.setEnabled(False)
        self.file_path_ctrl3.setStyleSheet("color: #3CB371;")
        self.file_path_ctrl3.setPlaceholderText("选择要录入的文件")
        self.file_path_choice_button3 = QPushButton("选择文件", self.widget3)
        self.file_path_choice_button3.clicked.connect(self.choice_file_path3)
        hbox_path.addWidget(self.file_path_ctrl3)
        hbox_path.addWidget(self.file_path_choice_button3)
        vbox2.addLayout(hbox_path)

        hbox_log = QHBoxLayout(self.widget3)
        self.log_ctrl3 = QTextEdit(self.widget3)
        self.log_ctrl3.setReadOnly(True)
        self.log_ctrl3.setStyleSheet("color: #3CB371;")
        hbox_log.addWidget(self.log_ctrl3)
        vbox2.addLayout(hbox_log)
        self.addTab(self.widget3, "CMS网站录入")


    def init_save(self):
        self.name_list = ["file_path_ctrl", "save_path_ctrl"]
        self.config_save.load_config(self.name_list)
        self.save_path_ctrl.textChanged.connect(self.save)
        self.file_path_ctrl.textChanged.connect(self.save)

    def save(self):
        self.config_save.save_config(self.name_list)

    def start_handle(self):
        """开始处理"""
        file_path = self.file_path_ctrl.text()
        if not os.path.exists(file_path):
            self.signal_hint_error.emit("要先选择一个有效的excel文件才能开始哦")
            return
        save_path = self.save_path_ctrl.text()
        if not os.path.exists(save_path):
            self.signal_hint_error.emit("保存路径不对诶")
            return
        thread = threading.Thread(target=self.file_handle.factory_excel_split, args=(file_path, save_path))
        thread.daemon = True
        thread.start()
        self.signal_log.emit("开始处理了哦")

    def choice_file_path(self):
        """加载url"""
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("EXCEL files (*.xls)")
        if file_dialog.exec_():

            selected_files = file_dialog.selectedFiles()
            file_path = selected_files[0]
            if not os.path.exists(file_path):
                self.signal_hint_error.emit("选择的这个文件不存在哦")
            else:
                self.file_path_ctrl.setText(file_path)

    def choice_save_path(self, message):
        """选择保存路径"""
        folder = QFileDialog.getExistingDirectory(self, '选择保存路径')
        if folder:
            self.save_path_ctrl.setText(folder)

    def start_handle2(self):
        """开始处理"""
        self.log_ctrl2.append("开始！")
        file_path = self.file_path_ctrl2.text()
        if not os.path.exists(file_path):
            self.signal_hint_error.emit("要先选择一个有效的文件夹才能开始哦")
            return
        save_path = self.save_path_ctrl2.text()
        if not os.path.exists(save_path):
            self.signal_hint_error.emit("保存路径不对诶")
            return
        value_list = []
        for file_name in os.listdir(file_path):
            if file_name.endswith(".jpg") or file_name.endswith(".png"):
                value_list.append(file_name.split(".")[0])
        content = "\n".join(value_list)
        with open(os.path.join(save_path, "output.txt"), "w", encoding="utf-8") as f:
            f.write(content)
        self.log_ctrl2.append("搞定！")

    def choice_file_path2(self):
        """加载url"""
        folder = QFileDialog.getExistingDirectory(self, '选择保存路径')
        if folder:
            self.file_path_ctrl2.setText(folder)

    def choice_save_path2(self):
        """选择保存路径"""
        folder = QFileDialog.getExistingDirectory(self, '选择保存路径')
        if folder:
            self.save_path_ctrl2.setText(folder)

    def choice_file_path3(self):
        """加载录入文件"""
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择XLSX文件",
            "",
            "Excel文件 (*.xlsx);;所有文件 (*)",
            options=options
        )

        if file_name:
            self.file_path_ctrl3.setText(file_name)
            # 解析excel
            self.extract_excel(file_name)

    def extract_excel(self, file_path):
        self.task_list = []
        try:
            df = pd.read_excel(file_path)
            name_list = list(df.iloc[:, 1])[1:]
            model_list = list(df.iloc[:, 2])[1:]
            color_list = list(df.iloc[:, 3])[1:]
            size_list = list(df.iloc[:, 4])[1:]
            material_list = list(df.iloc[:, 5])[1:]
            self.website_entry_info["Name"] = name_list
            self.website_entry_info["Model"] = model_list
            self.website_entry_info["Color"] = color_list
            self.website_entry_info["Size"] = size_list
            self.website_entry_info["Material"] = material_list
            self.log_ctrl3.append(f"解析表格成功，一共{len(name_list)}个！")
            return True
        except Exception as e:
            self.log_ctrl3.append(f"解析表格失败！原因：{e}")
            return False

    def open_website(self):
        """打开网站"""
        self.page = ChromiumPage()
        self.page.get("http://www.cnacczj.com/zy-manage/admin_login.php")
        self.log_ctrl3.append("网站打开了，需要自己登录一下哦")

    def website_entry(self):
        """网站录入"""
        try:
            body_node = self.page.ele("@id=frame_right").ele("tag=body")
        except Exception as e:
            self.log_ctrl3.append("页面不对哦，打开添加产品的页面再点录入")
            return
        if not self.is_not_none_node(body_node):
            self.log_ctrl3.append("页面不对哦，打开添加产品的页面再点录入")
            return
        try:
            current_num = int(self.current_row_ctrl.text())
            current_num = current_num - 3
            if current_num < 0:
                self.log_ctrl3.append("输入的行数要大于3哦")
                return
        except Exception as e:
            self.log_ctrl3.append("输入的行数不对哦")
            return
        name = self.website_entry_info["Name"][current_num]
        body_node.ele("@id=product_name").input(name)
        model = self.website_entry_info["Model"][current_num]
        body_node.ele("@id=product_model").input(model)
        color = self.website_entry_info["Color"][current_num]
        body_node.ele("@id=product_color").input(color)
        size = self.website_entry_info["Size"][current_num]
        body_node.ele("@id=product_size").input(size)
        material = self.website_entry_info["Material"][current_num]
        body_node.ele("@id=product_dis").input(material)
        self.log_ctrl3.append(f"把{self.current_row_ctrl.text()}填写进去了！^^")
        self.current_row_ctrl.setText(str(current_num + 4))

    def is_not_none_node(self, node):
        if node.__class__.__name__ != "NoneElement":
            return True
        else:
            return False

    def hint_error(self, message):
        """弹窗提示错误"""
        QMessageBox.information(self, "错误", message, QMessageBox.Yes)

    def hint_infomation(self, message):
        """弹窗提示信息"""
        QMessageBox.information(self, "提示", message, QMessageBox.Yes)

    def log(self, message):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        message = "--" + now + "--" + message
        self.log_ctrl.append(message)
        print(message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    frame = MainFrame()
    apply_stylesheet(app, theme="dark_lightgreen.xml")
    frame.show()
    sys.exit(app.exec_())