from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer, QUrl
from PyQt5.QtGui import QDesktopServices, QPixmap
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from image_ai.storage import ImageAiDatabase
from image_ai.task_manager import ImageAiTaskManager


IMAGE_FILTER = "图片文件 (*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff);;所有文件 (*)"


def _read_pixmap(path: Path) -> QPixmap:
    pixmap = QPixmap()
    try:
        pixmap.loadFromData(path.read_bytes())
    except OSError:
        return QPixmap()
    return pixmap


def _format_elapsed(milliseconds: int) -> str:
    seconds = max(0, milliseconds) // 1000
    if seconds < 60:
        return f"{seconds} 秒"
    minutes, seconds = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes} 分 {seconds} 秒"
    hours, minutes = divmod(minutes, 60)
    return f"{hours} 时 {minutes} 分"


def _task_elapsed_ms(task: dict[str, object]) -> int:
    status = str(task.get("status") or "")
    stored = max(0, int(task.get("elapsed_ms") or 0))
    if status not in {"queued", "running"}:
        return stored
    timestamp = task.get("started_at") if status == "running" else task.get("created_at")
    try:
        started = datetime.fromisoformat(str(timestamp or ""))
        now = datetime.now().astimezone()
        return max(0, round((now - started).total_seconds() * 1000))
    except (TypeError, ValueError):
        return stored


class ImagePreviewDialog(QDialog):
    def __init__(self, image_path: Path, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.image_path = image_path
        self.setWindowTitle(image_path.name)
        self.resize(900, 700)

        root = QVBoxLayout(self)
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        image_label = QLabel()
        image_label.setAlignment(Qt.AlignCenter)
        image_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        pixmap = _read_pixmap(image_path)
        if pixmap.isNull():
            image_label.setText("图片读取失败")
        else:
            image_label.setPixmap(
                pixmap.scaled(840, 610, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            )
        scroll.setWidget(image_label)
        root.addWidget(scroll, 1)

        path_label = QLabel(str(image_path))
        path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        path_label.setWordWrap(True)
        root.addWidget(path_label)

        actions = QHBoxLayout()
        open_folder_button = QPushButton("打开目录")
        save_as_button = QPushButton("另存为图片")
        close_button = QPushButton("关闭")
        open_folder_button.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(image_path.parent)))
        )
        save_as_button.clicked.connect(self.save_as)
        close_button.clicked.connect(self.accept)
        actions.addWidget(open_folder_button)
        actions.addWidget(save_as_button)
        actions.addStretch(1)
        actions.addWidget(close_button)
        root.addLayout(actions)

    def save_as(self) -> None:
        suffix = self.image_path.suffix.lower()
        image_filters = {
            ".png": "PNG 图片 (*.png)",
            ".jpg": "JPEG 图片 (*.jpg *.jpeg)",
            ".jpeg": "JPEG 图片 (*.jpg *.jpeg)",
            ".webp": "WebP 图片 (*.webp)",
        }
        selected_filter = image_filters.get(suffix, "图片文件 (*)")
        target_value, _filter = QFileDialog.getSaveFileName(
            self,
            "另存为图片",
            str(self.image_path.parent / self.image_path.name),
            f"{selected_filter};;所有文件 (*)",
        )
        if not target_value:
            return
        target = Path(target_value)
        if not target.suffix and suffix:
            target = target.with_suffix(suffix)
        try:
            if target.resolve() == self.image_path.resolve():
                QMessageBox.information(self, "另存为图片", "所选位置就是当前图片。")
                return
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(self.image_path, target)
        except OSError as exc:
            QMessageBox.warning(self, "另存为图片", f"保存图片失败：{exc}")
            return
        QMessageBox.information(self, "另存为图片", f"图片已保存到：\n{target}")


class ImageAiPage(QWidget):
    STATUS_TEXT = {
        "queued": "排队中",
        "running": "处理中",
        "completed": "已完成",
        "failed": "失败",
        "cancelled": "已取消",
        "interrupted": "已中断",
    }

    def __init__(
        self,
        database: ImageAiDatabase,
        manager: ImageAiTaskManager,
        project_root: Path,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.database = database
        self.manager = manager
        self.project_root = project_root
        self.source_path: Path | None = None
        self._build_ui()
        self._load_settings()

        self.manager.task_added.connect(self._task_changed)
        self.manager.task_updated.connect(self._task_changed)
        self.manager.active_count_changed.connect(self._update_summary)

        self.refresh_timer = QTimer(self)
        self.refresh_timer.setInterval(1000)
        self.refresh_timer.timeout.connect(self.refresh_tasks)
        self.refresh_timer.start()
        self.refresh_tasks()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        settings_group = QGroupBox("配置")
        settings_group.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        settings_layout = QHBoxLayout(settings_group)
        settings_layout.setContentsMargins(10, 7, 10, 7)
        settings_layout.setSpacing(8)

        settings_layout.addWidget(QLabel("API Key"))
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        self.api_key_edit.setPlaceholderText("请输入 API Key")
        self.show_api_key_check = QCheckBox("显示")
        self.save_api_key_button = QPushButton("保存")
        self.show_api_key_check.toggled.connect(self._toggle_api_key_visibility)
        self.save_api_key_button.clicked.connect(self.save_api_key)
        settings_layout.addWidget(self.api_key_edit, 3)
        settings_layout.addWidget(self.show_api_key_check)
        settings_layout.addWidget(self.save_api_key_button)

        settings_layout.addSpacing(8)
        settings_layout.addWidget(QLabel("结果目录"))
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        output_dir_button = QPushButton("选择目录")
        output_dir_button.clicked.connect(self.choose_output_dir)
        settings_layout.addWidget(self.output_dir_edit, 2)
        settings_layout.addWidget(output_dir_button)
        root.addWidget(settings_group)

        create_group = QGroupBox("新建任务")
        create_layout = QHBoxLayout(create_group)
        self.source_preview = QLabel("未选择图片\n文生图")
        self.source_preview.setAlignment(Qt.AlignCenter)
        self.source_preview.setFrameShape(QFrame.StyledPanel)
        self.source_preview.setFixedSize(150, 150)
        create_layout.addWidget(self.source_preview)

        input_layout = QVBoxLayout()
        source_layout = QHBoxLayout()
        self.source_path_edit = QLineEdit()
        self.source_path_edit.setReadOnly(True)
        self.source_path_edit.setPlaceholderText("不选图片时执行文生图")
        choose_image_button = QPushButton("选择图片")
        clear_image_button = QPushButton("清除")
        choose_image_button.clicked.connect(self.choose_source_image)
        clear_image_button.clicked.connect(self.clear_source_image)
        source_layout.addWidget(self.source_path_edit, 1)
        source_layout.addWidget(choose_image_button)
        source_layout.addWidget(clear_image_button)
        input_layout.addLayout(source_layout)

        self.prompt_edit = QPlainTextEdit()
        self.prompt_edit.setPlaceholderText("输入希望生成或修改的画面内容")
        self.prompt_edit.setMinimumHeight(82)
        input_layout.addWidget(self.prompt_edit, 1)

        submit_layout = QHBoxLayout()
        self.submit_status_label = QLabel()
        self.submit_button = QPushButton("加入队列")
        self.submit_button.setMinimumWidth(110)
        self.submit_button.clicked.connect(self.submit_task)
        submit_layout.addWidget(self.submit_status_label, 1)
        submit_layout.addWidget(self.submit_button)
        input_layout.addLayout(submit_layout)
        create_layout.addLayout(input_layout, 1)
        root.addWidget(create_group)

        tasks_header = QHBoxLayout()
        tasks_title = QLabel("任务记录")
        self.summary_label = QLabel()
        tasks_header.addWidget(tasks_title)
        tasks_header.addStretch(1)
        tasks_header.addWidget(self.summary_label)
        root.addLayout(tasks_header)

        self.task_table = QTableWidget(0, 7)
        self.task_table.setHorizontalHeaderLabels(
            ["编号", "类型", "源图片", "提示词", "状态", "耗时", "结果"]
        )
        self.task_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.task_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.task_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.task_table.setAlternatingRowColors(True)
        self.task_table.verticalHeader().setVisible(False)
        header = self.task_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.task_table.itemSelectionChanged.connect(self._update_task_actions)
        self.task_table.itemDoubleClicked.connect(lambda _item: self.view_result())
        root.addWidget(self.task_table, 1)

        task_actions = QHBoxLayout()
        self.view_result_button = QPushButton("查看结果")
        self.open_folder_button = QPushButton("打开目录")
        self.retry_button = QPushButton("重试")
        self.cancel_button = QPushButton("取消任务")
        clear_history_button = QPushButton("清空记录")
        refresh_button = QPushButton("刷新")
        self.view_result_button.clicked.connect(self.view_result)
        self.open_folder_button.clicked.connect(self.open_result_folder)
        self.retry_button.clicked.connect(self.retry_task)
        self.cancel_button.clicked.connect(self.cancel_task)
        clear_history_button.clicked.connect(self.clear_task_history)
        refresh_button.clicked.connect(self.refresh_tasks)
        task_actions.addWidget(self.view_result_button)
        task_actions.addWidget(self.open_folder_button)
        task_actions.addWidget(self.retry_button)
        task_actions.addWidget(self.cancel_button)
        task_actions.addStretch(1)
        task_actions.addWidget(clear_history_button)
        task_actions.addWidget(refresh_button)
        root.addLayout(task_actions)

    def _load_settings(self) -> None:
        default_output = self.project_root / "result" / "生图改图"
        output_dir = self.database.get_output_dir(str(default_output))
        self.output_dir_edit.setText(output_dir)
        try:
            self.api_key_edit.setText(self.database.get_api_key())
        except (OSError, ValueError) as exc:
            self.api_key_edit.clear()
            QTimer.singleShot(
                0,
                lambda message=str(exc): QMessageBox.warning(
                    self, "API Key", f"读取已保存的 API Key 失败：{message}"
                ),
            )

    def _toggle_api_key_visibility(self, visible: bool) -> None:
        self.api_key_edit.setEchoMode(QLineEdit.Normal if visible else QLineEdit.Password)

    def save_api_key(self, checked: bool = False, *, show_message: bool = True) -> bool:
        del checked
        try:
            self.database.set_api_key(self.api_key_edit.text())
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "API Key", f"保存 API Key 失败：{exc}")
            return False
        if show_message:
            message = "API Key 已加密保存。" if self.api_key_edit.text().strip() else "API Key 已清除。"
            QMessageBox.information(self, "API Key", message)
        return True

    def choose_output_dir(self) -> None:
        initial = self.output_dir_edit.text() or str(self.project_root)
        folder = QFileDialog.getExistingDirectory(self, "选择结果目录", initial)
        if not folder:
            return
        self.output_dir_edit.setText(folder)
        self.database.set_output_dir(folder)

    def choose_source_image(self) -> None:
        initial = str(self.source_path.parent) if self.source_path else str(self.project_root)
        file_path, _selected_filter = QFileDialog.getOpenFileName(
            self,
            "选择图片",
            initial,
            IMAGE_FILTER,
        )
        if file_path:
            self._set_source_image(Path(file_path))

    def clear_source_image(self) -> None:
        self.source_path = None
        self.source_path_edit.clear()
        self.source_preview.clear()
        self.source_preview.setText("未选择图片\n文生图")

    def _set_source_image(self, image_path: Path) -> None:
        pixmap = _read_pixmap(image_path)
        if pixmap.isNull():
            QMessageBox.warning(self, "选择图片", "无法读取选择的图片文件。")
            return
        self.source_path = image_path.resolve()
        self.source_path_edit.setText(str(self.source_path))
        self.source_preview.setText("")
        self.source_preview.setPixmap(
            pixmap.scaled(144, 144, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        )

    def submit_task(self) -> None:
        if not self.save_api_key(show_message=False):
            return
        try:
            task_id = self.manager.submit(
                api_key=self.api_key_edit.text(),
                prompt=self.prompt_edit.toPlainText(),
                output_dir=self.output_dir_edit.text(),
                source_path=self.source_path,
            )
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "生图改图", str(exc))
            return
        self.submit_status_label.setText(f"任务 #{task_id} 已加入队列")
        self.refresh_tasks()

    def _task_changed(self, _task_id: int) -> None:
        self.refresh_tasks()

    def _update_summary(self, _active_count: int | None = None) -> None:
        self.summary_label.setText(
            f"处理中 {self.manager.running_task_count()}/"
            f"{self.manager.MAX_CONCURRENT_TASKS}  ·  "
            f"排队 {self.manager.queued_task_count()}"
        )

    def refresh_tasks(self) -> None:
        selected_id = self._selected_task_id()
        tasks = self.database.list_tasks()
        self.task_table.setRowCount(len(tasks))
        selected_row = -1
        for row, task in enumerate(tasks):
            task_id = int(task.get("id") or 0)
            source_path = Path(str(task.get("source_path") or ""))
            output_path = Path(str(task.get("output_path") or ""))
            prompt = " ".join(str(task.get("prompt") or "").split())
            values = (
                str(task_id),
                "图片修改" if task.get("task_type") == "edit" else "文生图",
                source_path.name if str(task.get("source_path") or "") else "-",
                prompt,
                self.STATUS_TEXT.get(str(task.get("status") or ""), "未知"),
                _format_elapsed(_task_elapsed_ms(task)),
                output_path.name if str(task.get("output_path") or "") else "-",
            )
            error = str(task.get("error_message") or "").strip()
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setToolTip(error or (prompt if column == 3 else value))
                if column == 0:
                    item.setData(Qt.UserRole, task_id)
                self.task_table.setItem(row, column, item)
            if task_id == selected_id:
                selected_row = row
        if selected_row >= 0:
            self.task_table.selectRow(selected_row)
        self._update_summary()
        self._update_task_actions()

    def _selected_task_id(self) -> int:
        row = self.task_table.currentRow()
        item = self.task_table.item(row, 0) if row >= 0 else None
        return int(item.data(Qt.UserRole) or 0) if item is not None else 0

    def _selected_task(self) -> dict[str, object] | None:
        task_id = self._selected_task_id()
        return self.database.get_task(task_id) if task_id else None

    def _update_task_actions(self) -> None:
        task = self._selected_task()
        status = str(task.get("status") or "") if task else ""
        output_path = Path(str(task.get("output_path") or "")) if task else Path()
        output_dir = Path(str(task.get("output_dir") or "")) if task else Path()
        self.view_result_button.setEnabled(bool(task and output_path.is_file()))
        self.open_folder_button.setEnabled(bool(task and output_dir.is_dir()))
        self.retry_button.setEnabled(status in {"failed", "cancelled", "interrupted"})
        self.cancel_button.setEnabled(status in {"queued", "running"})

    def view_result(self) -> None:
        task = self._selected_task()
        if task is None:
            return
        output_path = Path(str(task.get("output_path") or ""))
        if not output_path.is_file():
            QMessageBox.warning(self, "查看结果", "该任务还没有可查看的结果图片。")
            return
        ImagePreviewDialog(output_path, self).exec_()

    def open_result_folder(self) -> None:
        task = self._selected_task()
        if task is None:
            return
        output_path = Path(str(task.get("output_path") or ""))
        output_dir = Path(str(task.get("output_dir") or ""))
        directory = output_path.parent if output_path.is_file() else output_dir
        if directory.is_dir():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(directory)))
        else:
            QMessageBox.warning(self, "打开目录", "任务结果目录不存在。")

    def retry_task(self) -> None:
        task_id = self._selected_task_id()
        if not task_id or not self.save_api_key(show_message=False):
            return
        try:
            new_task_id = self.manager.retry(task_id, self.api_key_edit.text())
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "重试任务", str(exc))
            return
        self.submit_status_label.setText(f"重试任务 #{new_task_id} 已加入队列")
        self.refresh_tasks()

    def cancel_task(self) -> None:
        task_id = self._selected_task_id()
        if task_id:
            self.manager.cancel_task(task_id)

    def clear_task_history(self) -> None:
        answer = QMessageBox.question(
            self,
            "清空任务记录",
            "确定清空任务记录吗？\n\n"
            "已生成的图片不会被删除，正在处理和排队的任务会保留。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        deleted_count = self.database.clear_task_history()
        self.refresh_tasks()
        QMessageBox.information(
            self,
            "清空任务记录",
            f"已清空 {deleted_count} 条任务记录。",
        )
