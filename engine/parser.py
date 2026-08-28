"""
e호조 23310 지출 엑셀/CSV 파일 파서 (e-Hojo 23310 Expense Parser)
- 유연한 헤더 매핑 (결의일자, 통계목, 지출액, 적요, 채권자)
- openpyxl 및 csv 파싱 지원
"""

import os
import re
from typing import List, Dict, Any, Tuple


class EHojoParser:
    # 컬럼 추정용 키워드 사전
    COL_MAPPINGS = {
        "date": ["결의일자", "원인행위일자", "지출일자", "집행일자", "일자", "작성일자"],
        "account": ["통계목", "예산과목", "목명", "편성목", "과목명", "세출과목"],
        "amount": ["지출액", "원인행위액", "지급액", "집행액", "결의금액", "금액", "지출금액"],
        "summary": ["적요", "원인행위적요", "지출건명", "결의적요", "건명", "사업내용", "내용"],
        "vendor": ["채권자", "지급처", "거래처", "수령인", "상호", "성명"],
        "code": ["결의번호", "원인행위번호", "지출번호", "문서번호", "번호"]
    }

    @classmethod
    def _clean_amount(cls, val: Any) -> int:
        if val is None:
            return 0
        if isinstance(val, (int, float)):
            return int(val)
        s = str(val).replace(",", "").replace("원", "").strip()
        try:
            return int(float(s))
        except ValueError:
            return 0

    @classmethod
    def _clean_str(cls, val: Any) -> str:
        if val is None:
            return ""
        return str(val).strip()

    @classmethod
    def parse_file(cls, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        """
        파일을 파싱하여 정규화된 지출 거래 목록을 반환합니다.
        반환값: (거래 리스트, 상태 메시지)
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        ext = os.path.splitext(file_path)[1].lower()
        if ext in (".xlsx", ".xlsm", ".xltx"):
            return cls._parse_excel(file_path)
        elif ext == ".csv":
            return cls._parse_csv(file_path)
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {ext} (.xlsx, .csv 지원)")

    @classmethod
    def _parse_excel(cls, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        try:
            import openpyxl
        except ImportError:
            raise ImportError("openpyxl 라이브러리가 필요합니다. pip install openpyxl")

        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active

        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return [], "엑셀 파일에 데이터가 없습니다."

        return cls._process_raw_rows(rows)

    @classmethod
    def _parse_csv(cls, file_path: str) -> Tuple[List[Dict[str, Any]], str]:
        import csv
        rows = []
        encodings = ["utf-8-sig", "euc-kr", "cp949", "utf-8"]
        for enc in encodings:
            try:
                with open(file_path, "r", encoding=enc) as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                break
            except (UnicodeDecodeError, Exception):
                continue
        
        if not rows:
            return [], "CSV 파일을 읽을 수 없거나 데이터가 비어있습니다."

        return cls._process_raw_rows(rows)

    @classmethod
    def _process_raw_rows(cls, rows: List[Tuple]) -> Tuple[List[Dict[str, Any]], str]:
        # 1. 헤더 행 탐색 (상위 15개 행 탐색)
        header_idx = -1
        col_indices = {}

        for r_idx, row in enumerate(rows[:15]):
            if not row:
                continue
            str_row = [cls._clean_str(cell) for cell in row]
            matched = {}
            for col_type, candidates in cls.COL_MAPPINGS.items():
                for c_idx, cell_text in enumerate(str_row):
                    if any(cand in cell_text for cand in candidates):
                        matched[col_type] = c_idx
                        break
            
            # 최소 지출액(amount)과 적요(summary) 컬럼이 매칭되면 헤더로 판단
            if "amount" in matched and ("summary" in matched or "account" in matched):
                header_idx = r_idx
                col_indices = matched
                break

        if header_idx == -1:
            # 헤더를 못 찾은 경우 기본 위치 추정 (0행 가정)
            header_idx = 0
            for c_idx, cell in enumerate(rows[0]):
                txt = cls._clean_str(cell)
                for col_type, candidates in cls.COL_MAPPINGS.items():
                    if any(cand in txt for cand in candidates) and col_type not in col_indices:
                        col_indices[col_type] = c_idx

        # 2. 데이터 행 파싱
        results = []
        for row in rows[header_idx + 1:]:
            if not row:
                continue
            
            amount = 0
            if "amount" in col_indices and col_indices["amount"] < len(row):
                amount = cls._clean_amount(row[col_indices["amount"]])
            
            # 금액이 0이거나 빈 행 건너뜀 (합계행 등 제외)
            if amount == 0:
                continue

            date_val = ""
            if "date" in col_indices and col_indices["date"] < len(row):
                date_val = cls._clean_str(row[col_indices["date"]])

            account_val = ""
            if "account" in col_indices and col_indices["account"] < len(row):
                account_val = cls._clean_str(row[col_indices["account"]])

            summary_val = ""
            if "summary" in col_indices and col_indices["summary"] < len(row):
                summary_val = cls._clean_str(row[col_indices["summary"]])

            vendor_val = ""
            if "vendor" in col_indices and col_indices["vendor"] < len(row):
                vendor_val = cls._clean_str(row[col_indices["vendor"]])

            code_val = ""
            if "code" in col_indices and col_indices["code"] < len(row):
                code_val = cls._clean_str(row[col_indices["code"]])

            # '합계', '소계', '총계' 행 제외
            if any(term in summary_val for term in ["합계", "소계", "총계", "누계"]):
                continue

            results.append({
                "code": code_val,
                "date": date_val,
                "account": account_val,
                "sub_account": "미분류",
                "summary": summary_val,
                "vendor": vendor_val,
                "amount": amount,
                "rule_matched": ""
            })

        msg = f"총 {len(results)}건의 지출 내역을 성공적으로 추출했습니다."
        return results, msg
