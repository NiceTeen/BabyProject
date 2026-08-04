from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

from image_ai.credentials import protect_secret, unprotect_secret


class ImageAiDatabase:
    API_KEY_SETTING = "image_ai_api_key"
    OUTPUT_DIR_SETTING = "image_ai_output_dir"

    def __init__(self, path: Path | str) -> None:
        self.path = Path(path)
        self.initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(str(self.path), timeout=15)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=WAL")
        return connection

    def initialize(self) -> None:
        with self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS config (
                    Name TEXT PRIMARY KEY,
                    Value TEXT
                );

                CREATE TABLE IF NOT EXISTS image_ai_tasks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_type TEXT NOT NULL,
                    source_path TEXT NOT NULL DEFAULT '',
                    output_dir TEXT NOT NULL,
                    output_path TEXT NOT NULL DEFAULT '',
                    prompt TEXT NOT NULL,
                    model TEXT NOT NULL,
                    base_url TEXT NOT NULL,
                    status TEXT NOT NULL,
                    error_message TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL,
                    started_at TEXT NOT NULL DEFAULT '',
                    finished_at TEXT NOT NULL DEFAULT '',
                    elapsed_ms INTEGER NOT NULL DEFAULT 0
                );

                CREATE INDEX IF NOT EXISTS idx_image_ai_tasks_created_at
                ON image_ai_tasks(created_at DESC, id DESC);
                """
            )
            now = datetime.now().astimezone().isoformat(timespec="seconds")
            connection.execute(
                """
                UPDATE image_ai_tasks
                SET status = 'interrupted',
                    error_message = CASE
                        WHEN error_message = '' THEN '软件上次退出时任务尚未完成。'
                        ELSE error_message
                    END,
                    finished_at = CASE WHEN finished_at = '' THEN ? ELSE finished_at END
                WHERE status IN ('queued', 'running')
                """,
                (now,),
            )

    def get_setting(self, name: str, default: str = "") -> str:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT Value FROM config WHERE Name = ?", (name,)
            ).fetchone()
        if row is None:
            return default
        try:
            value = json.loads(str(row["Value"]))
        except (TypeError, ValueError):
            return default
        return str(value) if value is not None else default

    def set_setting(self, name: str, value: str) -> None:
        content = json.dumps(str(value), ensure_ascii=False)
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO config(Name, Value) VALUES (?, ?)
                ON CONFLICT(Name) DO UPDATE SET Value = excluded.Value
                """,
                (name, content),
            )

    def get_api_key(self) -> str:
        protected = self.get_setting(self.API_KEY_SETTING)
        return unprotect_secret(protected) if protected else ""

    def set_api_key(self, api_key: str) -> None:
        self.set_setting(self.API_KEY_SETTING, protect_secret(api_key.strip()))

    def get_output_dir(self, default: str) -> str:
        return self.get_setting(self.OUTPUT_DIR_SETTING, default)

    def set_output_dir(self, output_dir: str) -> None:
        self.set_setting(self.OUTPUT_DIR_SETTING, output_dir)

    def add_task(self, values: dict[str, Any]) -> int:
        columns = (
            "task_type",
            "source_path",
            "output_dir",
            "output_path",
            "prompt",
            "model",
            "base_url",
            "status",
            "error_message",
            "created_at",
            "started_at",
            "finished_at",
            "elapsed_ms",
        )
        with self._connect() as connection:
            cursor = connection.execute(
                f"INSERT INTO image_ai_tasks({', '.join(columns)}) "
                f"VALUES ({', '.join('?' for _ in columns)})",
                tuple(values.get(column, "") for column in columns),
            )
            return int(cursor.lastrowid)

    def get_task(self, task_id: int) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM image_ai_tasks WHERE id = ?", (task_id,)
            ).fetchone()
        return dict(row) if row is not None else None

    def list_tasks(self, limit: int = 300) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT * FROM image_ai_tasks "
                "ORDER BY created_at DESC, id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [dict(row) for row in rows]

    def update_task(self, task_id: int, **values: Any) -> None:
        allowed = {
            "output_path",
            "status",
            "error_message",
            "started_at",
            "finished_at",
            "elapsed_ms",
        }
        updates = [(key, value) for key, value in values.items() if key in allowed]
        if not updates:
            return
        assignments = ", ".join(f"{key} = ?" for key, _value in updates)
        parameters = [value for _key, value in updates]
        parameters.append(task_id)
        with self._connect() as connection:
            connection.execute(
                f"UPDATE image_ai_tasks SET {assignments} WHERE id = ?", parameters
            )

    def clear_task_history(self) -> int:
        """Delete inactive task records without removing generated image files."""
        with self._connect() as connection:
            cursor = connection.execute(
                "DELETE FROM image_ai_tasks WHERE status NOT IN ('queued', 'running')"
            )
            return max(0, cursor.rowcount)
