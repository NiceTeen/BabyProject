# -*- coding:utf-8 -*-
import json
import os
import sys
from datetime import datetime
from pathlib import Path
import sqlite3
import threading
import xlrd
import xlwt
import pandas as pd

from DrissionPage import ChromiumPage, ChromiumOptions
from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIntValidator
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from image_ai import ImageAiDatabase, ImageAiPage, ImageAiTaskManager
from ui_theme import apply_app_theme


ChromiumOptions().set_browser_path(
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
).save()


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
        self.image_ai_database = ImageAiDatabase(Path("Config.db"))
        self.image_ai_manager = ImageAiTaskManager(self.image_ai_database, self)
        self.website_entry_info = {}
        self.init_ui()
        self.log("日志系统")
        self.init_save()

        # 信号绑定
        self.signal_log.connect(self.log)
        self.signal_hint_info.connect(self.hint_infomation)
        self.signal_hint_error.connect(self.hint_error)


    @staticmethod
    def _set_role(widget, name, value):
        widget.setProperty(name, value)
        return widget

    def _create_page(self, eyebrow, title, description, badge):
        page = QWidget(self)
        page.setObjectName("AppPage")
        root = QVBoxLayout(page)
        root.setContentsMargins(24, 22, 24, 24)
        root.setSpacing(16)

        header = QFrame(page)
        header.setObjectName("PageHeader")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(22, 17, 20, 17)
        header_layout.setSpacing(18)

        copy_layout = QVBoxLayout()
        copy_layout.setSpacing(3)
        eyebrow_label = QLabel(eyebrow, header)
        self._set_role(eyebrow_label, "uiRole", "eyebrow")
        title_label = QLabel(title, header)
        self._set_role(title_label, "uiRole", "pageTitle")
        description_label = QLabel(description, header)
        description_label.setWordWrap(True)
        self._set_role(description_label, "uiRole", "pageDescription")
        copy_layout.addWidget(eyebrow_label)
        copy_layout.addWidget(title_label)
        copy_layout.addWidget(description_label)
        header_layout.addLayout(copy_layout, 1)

        badge_label = QLabel(badge, header)
        badge_label.setAlignment(Qt.AlignCenter)
        self._set_role(badge_label, "uiRole", "badge")
        header_layout.addWidget(badge_label, 0, Qt.AlignTop)
        root.addWidget(header)
        return page, root

    def _create_card(self, parent, title, hint=""):
        card = QFrame(parent)
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 17, 20, 19)
        layout.setSpacing(12)

        title_row = QHBoxLayout()
        title_label = QLabel(title, card)
        self._set_role(title_label, "uiRole", "sectionTitle")
        title_row.addWidget(title_label)
        if hint:
            title_row.addStretch(1)
            hint_label = QLabel(hint, card)
            self._set_role(hint_label, "uiRole", "hint")
            title_row.addWidget(hint_label)
        layout.addLayout(title_row)
        return card, layout

    def _add_path_row(self, layout, parent, label_text, control, button):
        row = QHBoxLayout()
        row.setSpacing(10)
        label = QLabel(label_text, parent)
        label.setFixedWidth(72)
        self._set_role(label, "uiRole", "fieldLabel")
        row.addWidget(label)
        row.addWidget(control, 1)
        button.setMinimumWidth(100)
        row.addWidget(button)
        layout.addLayout(row)

    def _add_log_card(self, root, parent, log_control, hint="运行信息会记录在这里"):
        card, layout = self._create_card(parent, "处理记录", hint)
        log_control.setReadOnly(True)
        log_control.setObjectName("LogPanel")
        layout.addWidget(log_control, 1)
        root.addWidget(card, 1)

    def init_ui(self):
        self.setWindowTitle(" ")
        self.setWindowIcon(QtGui.QIcon("ico.ico"))
        self.resize(1120, 790)
        self.setMinimumSize(960, 680)
        self.setDocumentMode(True)
        self.setTabPosition(QTabWidget.North)

        # 拆分功能
        self.widget1, page1_layout = self._create_page(
            "EXCEL TOOLS",
            "报价表拆分",
            "按工厂自动整理并生成独立表格，常用路径会继续为你保留。",
            "省心整理",
        )
        split_card, split_layout = self._create_card(
            self.widget1, "文件设置", "支持 .xls 格式"
        )
        self.file_path_ctrl = QLineEdit(self.widget1)
        self.file_path_ctrl.setReadOnly(True)
        self.file_path_ctrl.setPlaceholderText("请选择需要拆分的 Excel 文件")
        self.file_path_choice_button = QPushButton("选择文件", self.widget1)
        self.file_path_choice_button.clicked.connect(self.choice_file_path)
        self._add_path_row(
            split_layout,
            split_card,
            "源文件",
            self.file_path_ctrl,
            self.file_path_choice_button,
        )

        self.save_path_ctrl = QLineEdit(self.widget1)
        self.save_path_ctrl.setReadOnly(True)
        self.save_path_ctrl.setPlaceholderText("请选择拆分后文件的存放目录")
        self.save_path_choice_button = QPushButton("选择目录", self.widget1)
        self.save_path_choice_button.clicked.connect(self.choice_save_path)
        self._add_path_row(
            split_layout,
            split_card,
            "保存到",
            self.save_path_ctrl,
            self.save_path_choice_button,
        )

        split_actions = QHBoxLayout()
        split_tip = QLabel("确认文件与目录后即可开始，处理期间可以查看下方记录。", split_card)
        self._set_role(split_tip, "uiRole", "hint")
        self.start_button = QPushButton("开始拆分", self.widget1)
        self._set_role(self.start_button, "buttonRole", "primary")
        self.start_button.setMinimumWidth(128)
        self.start_button.clicked.connect(self.start_handle)
        split_actions.addWidget(split_tip)
        split_actions.addStretch(1)
        split_actions.addWidget(self.start_button)
        split_layout.addLayout(split_actions)
        page1_layout.addWidget(split_card)

        self.log_ctrl = QTextEdit(self.widget1)
        self._add_log_card(page1_layout, self.widget1, self.log_ctrl)
        self.addTab(self.widget1, "表格拆分")

        # 提取货号
        self.widget2, page2_layout = self._create_page(
            "IMAGE TOOLS",
            "图片货号提取",
            "读取文件夹内的图片名称，快速汇总为可复制使用的货号清单。",
            "快速清单",
        )
        extract_card, extract_layout = self._create_card(
            self.widget2, "文件夹设置", "识别 .jpg 与 .png 图片"
        )
        self.file_path_ctrl2 = QLineEdit(self.widget2)
        self.file_path_ctrl2.setReadOnly(True)
        self.file_path_ctrl2.setPlaceholderText("请选择存放图片的文件夹")
        self.file_path_choice_button2 = QPushButton("选择文件夹", self.widget2)
        self.file_path_choice_button2.clicked.connect(self.choice_file_path2)
        self._add_path_row(
            extract_layout,
            extract_card,
            "图片目录",
            self.file_path_ctrl2,
            self.file_path_choice_button2,
        )

        self.save_path_ctrl2 = QLineEdit(self.widget2)
        self.save_path_ctrl2.setReadOnly(True)
        self.save_path_ctrl2.setPlaceholderText("请选择 output.txt 的存放目录")
        self.save_path_choice_button2 = QPushButton("选择目录", self.widget2)
        self.save_path_choice_button2.clicked.connect(self.choice_save_path2)
        self._add_path_row(
            extract_layout,
            extract_card,
            "保存到",
            self.save_path_ctrl2,
            self.save_path_choice_button2,
        )

        extract_actions = QHBoxLayout()
        extract_tip = QLabel("结果会保存为 output.txt，每个货号单独一行。", extract_card)
        self._set_role(extract_tip, "uiRole", "hint")
        self.start_button2 = QPushButton("开始提取", self.widget2)
        self._set_role(self.start_button2, "buttonRole", "primary")
        self.start_button2.setMinimumWidth(128)
        self.start_button2.clicked.connect(self.start_handle2)
        extract_actions.addWidget(extract_tip)
        extract_actions.addStretch(1)
        extract_actions.addWidget(self.start_button2)
        extract_layout.addLayout(extract_actions)
        page2_layout.addWidget(extract_card)

        self.log_ctrl2 = QTextEdit(self.widget2)
        self._add_log_card(page2_layout, self.widget2, self.log_ctrl2)
        self.addTab(self.widget2, "货号提取")

        # 网站录入
        self.widget3, page3_layout = self._create_page(
            "CMS ASSISTANT",
            "CMS 网站录入",
            "从 Excel 读取产品资料，按行辅助填写网站表单，减少重复输入。",
            "录入助手",
        )
        entry_card, entry_layout = self._create_card(
            self.widget3, "录入准备", "建议先选择表格，再打开网站登录"
        )
        self.file_path_ctrl3 = QLineEdit(self.widget3)
        self.file_path_ctrl3.setReadOnly(True)
        self.file_path_ctrl3.setPlaceholderText("请选择需要录入的 .xlsx 文件")
        self.file_path_choice_button3 = QPushButton("选择文件", self.widget3)
        self.file_path_choice_button3.clicked.connect(self.choice_file_path3)
        self._add_path_row(
            entry_layout,
            entry_card,
            "数据表格",
            self.file_path_ctrl3,
            self.file_path_choice_button3,
        )

        entry_actions = QHBoxLayout()
        entry_actions.setSpacing(10)
        row_label = QLabel("当前行", entry_card)
        self._set_role(row_label, "uiRole", "fieldLabel")
        self.current_row_ctrl = QLineEdit(self.widget3)
        self.current_row_ctrl.setText("3")
        self.current_row_ctrl.setValidator(QIntValidator(1, 999999, self))
        self.current_row_ctrl.setAlignment(Qt.AlignCenter)
        self.current_row_ctrl.setFixedWidth(78)
        self.current_row_ctrl.setToolTip("Excel 中准备录入的数据行号")
        self.open_website_button = QPushButton("打开 CMS 网站", self.widget3)
        self.open_website_button.clicked.connect(self.open_website)
        self.entry_button = QPushButton("录入这一行", self.widget3)
        self._set_role(self.entry_button, "buttonRole", "primary")
        self.entry_button.clicked.connect(self.website_entry)
        entry_actions.addWidget(row_label)
        entry_actions.addWidget(self.current_row_ctrl)
        entry_actions.addSpacing(8)
        entry_actions.addWidget(self.open_website_button)
        entry_actions.addStretch(1)
        entry_actions.addWidget(self.entry_button)
        entry_layout.addLayout(entry_actions)
        page3_layout.addWidget(entry_card)

        self.log_ctrl3 = QTextEdit(self.widget3)
        self._add_log_card(
            page3_layout,
            self.widget3,
            self.log_ctrl3,
            "表格解析与网站录入状态会显示在这里",
        )
        self.addTab(self.widget3, "CMS 录入")

        self.image_ai_page = ImageAiPage(
            self.image_ai_database,
            self.image_ai_manager,
            Path.cwd(),
            self,
        )
        self.addTab(self.image_ai_page, "生图改图")


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
            subclass_list = list(df.iloc[:, 7])[1:]
            series_list = list(df.iloc[:, 8])[1:]
            self.website_entry_info["Name"] = name_list
            self.website_entry_info["Model"] = model_list
            self.website_entry_info["Color"] = color_list
            self.website_entry_info["Size"] = size_list
            self.website_entry_info["Material"] = material_list
            self.website_entry_info["SubClass"] = subclass_list
            self.website_entry_info["Series"] = series_list
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
        except Exception:
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
        except Exception:
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
        subclass = self.website_entry_info["SubClass"][current_num]
        subclass_select_node = body_node.ele("@id=subclass_id")
        for option_node in subclass_select_node.eles("tag=option"):
            if subclass.upper() == option_node.text.replace("|- ", "").upper():
                subclass_select_node.select.by_value(option_node.value)
                break
        series = self.website_entry_info["Series"][current_num]
        series_select_node = body_node.ele("@id=collclass_id")
        for option_node in series_select_node.eles("tag=option"):
            if series.upper() == option_node.text.replace("|- ", "").upper():
                series_select_node.select.by_value(option_node.value)
                break
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

    def closeEvent(self, event):
        if self.image_ai_manager.has_active_tasks():
            answer = QMessageBox.question(
                self,
                "图片 AI 任务进行中",
                "当前仍有生图改图任务。是否取消全部任务？后台任务结束后可再次关闭软件。",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer == QMessageBox.Yes:
                self.image_ai_manager.cancel_all()
            event.ignore()
            return
        self.image_ai_manager.shutdown()
        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_app_theme(app)
    frame = MainFrame()
    frame.show()
    sys.exit(app.exec_())
