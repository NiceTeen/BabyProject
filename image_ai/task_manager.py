from __future__ import annotations

import logging
import time
from collections import deque
from concurrent.futures import Future, ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from threading import Event

from PyQt5.QtCore import QObject, Qt, pyqtSignal, pyqtSlot

from image_ai.client import (
    IMAGE_AI_BASE_URL,
    IMAGE_AI_MODEL,
    ImageAiCancelled,
    edit_image,
    generate_image,
    save_generated_image,
    validate_api_key,
)
from image_ai.storage import ImageAiDatabase


logger = logging.getLogger(__name__)


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


@dataclass(frozen=True)
class ImageAiTaskResult:
    status: str
    elapsed_ms: int
    output_path: str = ""
    error_message: str = ""


def _safe_error(exc: Exception, api_key: str) -> str:
    message = str(exc) or exc.__class__.__name__
    if api_key:
        message = message.replace(api_key, "[已隐藏 API Key]")
    return message[:4000]


def _run_image_ai_task(
    task_id: int,
    api_key: str,
    source_path: Path | None,
    output_dir: Path,
    prompt: str,
    cancel_event: Event,
) -> ImageAiTaskResult:
    started = time.perf_counter()
    try:
        if source_path is None:
            generated = generate_image(
                api_key,
                prompt,
                cancel_event=cancel_event,
            )
        else:
            generated = edit_image(
                api_key,
                source_path,
                prompt,
                cancel_event=cancel_event,
            )
        if cancel_event.is_set():
            raise ImageAiCancelled("图片 AI 任务已取消。")
        output_path = save_generated_image(
            generated,
            output_dir,
            task_id,
            source_path,
        )
    except ImageAiCancelled:
        return ImageAiTaskResult(
            "cancelled", round((time.perf_counter() - started) * 1000)
        )
    except Exception as exc:
        logger.exception("图片 AI 任务失败：task_id=%s", task_id)
        return ImageAiTaskResult(
            "failed",
            round((time.perf_counter() - started) * 1000),
            error_message=_safe_error(exc, api_key),
        )
    return ImageAiTaskResult(
        "completed",
        round((time.perf_counter() - started) * 1000),
        output_path=str(output_path),
    )


class ImageAiTaskManager(QObject):
    task_added = pyqtSignal(int)
    task_updated = pyqtSignal(int)
    image_created = pyqtSignal(int, str)
    active_count_changed = pyqtSignal(int)
    _future_ready = pyqtSignal(int, object)

    MAX_CONCURRENT_TASKS = 3

    def __init__(
        self,
        database: ImageAiDatabase,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self.database = database
        self._queued: deque[int] = deque()
        self._api_keys: dict[int, str] = {}
        self._executor = ThreadPoolExecutor(
            max_workers=self.MAX_CONCURRENT_TASKS,
            thread_name_prefix="image-ai",
        )
        self._futures: dict[int, Future[ImageAiTaskResult]] = {}
        self._cancel_events: dict[int, Event] = {}
        self._future_ready.connect(self._future_finished, Qt.QueuedConnection)

    def submit(
        self,
        *,
        api_key: str,
        prompt: str,
        output_dir: Path | str,
        source_path: Path | str | None = None,
    ) -> int:
        api_key = validate_api_key(api_key)
        prompt = str(prompt).strip()
        if not prompt:
            raise ValueError("请输入生图提示词。")

        source = Path(source_path).resolve() if source_path else None
        if source is not None and not source.is_file():
            raise ValueError("选择的源图片不存在。")
        output = Path(output_dir).expanduser().resolve()
        try:
            output.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise ValueError(f"无法创建结果目录：{exc}") from exc

        task_type = "edit" if source is not None else "generate"
        task_id = self.database.add_task(
            {
                "task_type": task_type,
                "source_path": str(source) if source else "",
                "output_dir": str(output),
                "output_path": "",
                "prompt": prompt,
                "model": IMAGE_AI_MODEL,
                "base_url": IMAGE_AI_BASE_URL,
                "status": "queued",
                "error_message": "",
                "created_at": _now_iso(),
                "started_at": "",
                "finished_at": "",
                "elapsed_ms": 0,
            }
        )
        self._api_keys[task_id] = api_key
        self._queued.append(task_id)
        self.task_added.emit(task_id)
        self._emit_active_count()
        self._pump_queue()
        return task_id

    def retry(self, task_id: int, api_key: str) -> int:
        task = self.database.get_task(task_id)
        if task is None:
            raise ValueError("找不到需要重试的图片 AI 任务。")
        source_path = str(task.get("source_path") or "") or None
        return self.submit(
            api_key=api_key,
            prompt=str(task.get("prompt") or ""),
            output_dir=str(task.get("output_dir") or ""),
            source_path=source_path,
        )

    def has_active_tasks(self) -> bool:
        return bool(self._queued or self._futures)

    def active_task_count(self) -> int:
        return len(self._queued) + len(self._futures)

    def queued_task_count(self) -> int:
        return len(self._queued)

    def running_task_count(self) -> int:
        return len(self._futures)

    def cancel_task(self, task_id: int) -> bool:
        if task_id in self._queued:
            self._queued.remove(task_id)
            self._api_keys.pop(task_id, None)
            self.database.update_task(
                task_id,
                status="cancelled",
                error_message="任务已取消。",
                finished_at=_now_iso(),
                elapsed_ms=0,
            )
            self.task_updated.emit(task_id)
            self._emit_active_count()
            return True
        cancel_event = self._cancel_events.get(task_id)
        if cancel_event is not None:
            cancel_event.set()
            return True
        return False

    def cancel_all(self) -> None:
        for task_id in list(self._queued):
            self.cancel_task(task_id)
        for cancel_event in self._cancel_events.values():
            cancel_event.set()

    def shutdown(self) -> None:
        self.cancel_all()
        self._executor.shutdown(wait=False, cancel_futures=True)

    def _pump_queue(self) -> None:
        while self._queued and len(self._futures) < self.MAX_CONCURRENT_TASKS:
            task_id = self._queued.popleft()
            task = self.database.get_task(task_id)
            api_key = self._api_keys.get(task_id)
            if task is None or api_key is None:
                continue
            self._start_task(task_id, task, api_key)

    def _start_task(
        self,
        task_id: int,
        task: dict[str, object],
        api_key: str,
    ) -> None:
        cancel_event = Event()
        self._cancel_events[task_id] = cancel_event
        self.database.update_task(
            task_id,
            status="running",
            error_message="",
            started_at=_now_iso(),
        )
        self.task_updated.emit(task_id)
        source_value = str(task.get("source_path") or "")
        source_path = Path(source_value) if source_value else None
        try:
            future = self._executor.submit(
                _run_image_ai_task,
                task_id,
                api_key,
                source_path,
                Path(str(task.get("output_dir") or "")),
                str(task.get("prompt") or ""),
                cancel_event,
            )
        except RuntimeError as exc:
            self._cancel_events.pop(task_id, None)
            self._api_keys.pop(task_id, None)
            self._task_failed(task_id, f"无法启动图片 AI 任务：{exc}", 0)
            self._emit_active_count()
            return
        self._futures[task_id] = future
        self._emit_active_count()
        future.add_done_callback(
            lambda current_future, current_id=task_id: self._future_ready.emit(
                current_id, current_future
            )
        )

    @pyqtSlot(int, object)
    def _future_finished(
        self,
        task_id: int,
        future: Future[ImageAiTaskResult],
    ) -> None:
        if self._futures.get(task_id) is not future:
            return
        try:
            result = future.result()
        except Exception as exc:
            logger.exception("图片 AI 工作线程异常退出：task_id=%s", task_id)
            result = ImageAiTaskResult(
                "failed",
                0,
                error_message=str(exc) or exc.__class__.__name__,
            )

        if result.status == "completed":
            self._task_completed(task_id, result.output_path, result.elapsed_ms)
        elif result.status == "cancelled":
            self._task_cancelled(task_id, result.elapsed_ms)
        else:
            self._task_failed(task_id, result.error_message, result.elapsed_ms)
        self._task_finished(task_id)

    def _task_completed(self, task_id: int, output_path: str, elapsed_ms: int) -> None:
        self.database.update_task(
            task_id,
            output_path=output_path,
            status="completed",
            error_message="",
            finished_at=_now_iso(),
            elapsed_ms=max(0, elapsed_ms),
        )
        self.task_updated.emit(task_id)
        self.image_created.emit(task_id, output_path)

    def _task_failed(self, task_id: int, error: str, elapsed_ms: int) -> None:
        self.database.update_task(
            task_id,
            status="failed",
            error_message=str(error)[:4000],
            finished_at=_now_iso(),
            elapsed_ms=max(0, elapsed_ms),
        )
        self.task_updated.emit(task_id)

    def _task_cancelled(self, task_id: int, elapsed_ms: int) -> None:
        self.database.update_task(
            task_id,
            status="cancelled",
            error_message="任务已取消。",
            finished_at=_now_iso(),
            elapsed_ms=max(0, elapsed_ms),
        )
        self.task_updated.emit(task_id)

    def _task_finished(self, task_id: int) -> None:
        self._futures.pop(task_id, None)
        self._cancel_events.pop(task_id, None)
        self._api_keys.pop(task_id, None)
        self._emit_active_count()
        self._pump_queue()

    def _emit_active_count(self) -> None:
        self.active_count_changed.emit(self.active_task_count())
