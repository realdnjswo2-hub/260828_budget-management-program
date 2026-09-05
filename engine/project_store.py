"""Portable single-file storage for the budget management application."""

import json
import os
import tempfile
from datetime import datetime


class ProjectStore:
    MAGIC = "EHOJO_BUDGET_PROJECT"
    FORMAT_VERSION = 1
    EXTENSION = ".ebudget"

    LIST_FIELDS = (
        "budget_master",
        "supplementary_budgets",
        "rules",
        "recurring_plans",
        "scheduled_plans",
        "raw_transactions",
        "transactions",
        "expense_sources",
    )

    @classmethod
    def save(cls, file_path, state):
        if not file_path:
            raise ValueError("저장할 파일 경로가 없습니다.")
        if not file_path.lower().endswith(cls.EXTENSION):
            file_path += cls.EXTENSION

        destination = os.path.abspath(file_path)
        parent = os.path.dirname(destination)
        os.makedirs(parent, exist_ok=True)

        payload = {
            "format": cls.MAGIC,
            "format_version": cls.FORMAT_VERSION,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "data": dict(state),
        }

        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                suffix=".tmp",
                prefix=".ebudget_",
                dir=parent,
                delete=False,
            ) as stream:
                temp_path = stream.name
                json.dump(payload, stream, ensure_ascii=False, indent=2)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temp_path, destination)
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
        return destination

    @classmethod
    def load(cls, file_path):
        with open(file_path, "r", encoding="utf-8") as stream:
            payload = json.load(stream)

        if not isinstance(payload, dict) or payload.get("format") != cls.MAGIC:
            raise ValueError("예산관리대장 저장 파일 형식이 아닙니다.")
        if int(payload.get("format_version", 0)) > cls.FORMAT_VERSION:
            raise ValueError("현재 프로그램보다 새로운 버전에서 저장된 파일입니다.")

        state = payload.get("data")
        if not isinstance(state, dict):
            raise ValueError("저장 파일의 데이터가 손상되었습니다.")
        for field in cls.LIST_FIELDS:
            value = state.get(field, [])
            if not isinstance(value, list):
                raise ValueError(f"저장 파일의 '{field}' 항목이 올바르지 않습니다.")
            state[field] = value
        return state
