"""
e호조 지출 파일 파서 (e-Hojo Expense Parser)
- 지출 집행현황 PDF 표 파싱 (사업명, 통계목, 적요, 결의금액, 지급일자)
- 기존 23310 엑셀/CSV 형식 호환
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
        elif ext == ".pdf":
            transactions, metadata, message = cls.parse_expense_pdf(file_path)
            return transactions, message
        else:
            raise ValueError(f"지원하지 않는 파일 형식입니다: {ext} (.pdf, .xlsx, .csv 지원)")

    @classmethod
    def parse_expense_pdf(cls, file_path: str):
        """e호조 '지출 집행현황 조회' PDF 한 개를 파싱한다."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError("PDF 분석에 pdfplumber 라이브러리가 필요합니다.") from exc

        transactions = []
        year = None
        department = ""

        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                raise ValueError("PDF에 읽을 수 있는 페이지가 없습니다.")

            first_text = pdf.pages[0].extract_text() or ""
            if "지출" not in first_text or "집행현황" not in first_text:
                raise ValueError("e호조 '지출 집행현황 조회' PDF 형식을 확인할 수 없습니다.")

            year_match = re.search(r"(20\d{2})\s*년도\s*지출\s*집행현황", first_text)
            if year_match:
                year = int(year_match.group(1))

            department_match = re.search(r"부\s*서\s*:\s*([^\n]+?)(?:\s+출력일자\s*:|\n)", first_text)
            if department_match:
                department = cls._clean_pdf_text(department_match.group(1))

            for page_number, page in enumerate(pdf.pages, start=1):
                for table in page.extract_tables() or []:
                    transactions.extend(
                        cls._transactions_from_expense_table(
                            table,
                            source_file=os.path.basename(file_path),
                            page_number=page_number,
                            department=department,
                        )
                    )

        transactions = cls.deduplicate_transactions(transactions)
        if not transactions:
            raise ValueError(
                "PDF 표에서 결의금액이 있는 지출 내역을 찾지 못했습니다. "
                "e호조의 '지출 집행현황 조회' 출력물인지 확인하세요."
            )

        detail_projects = sorted({tx.get("detail_project", "") for tx in transactions if tx.get("detail_project")})
        metadata = {
            "file_name": os.path.basename(file_path),
            "year": year,
            "department": department,
            "detail_projects": detail_projects,
            "transaction_count": len(transactions),
            "total_amount": sum(int(tx.get("amount", 0)) for tx in transactions),
        }
        detail_label = ", ".join(detail_projects) if detail_projects else "사업명 미확인"
        message = (
            f"{os.path.basename(file_path)}: {detail_label} 지출 "
            f"{len(transactions)}건, {metadata['total_amount']:,}원 추출"
        )
        return transactions, metadata, message

    @classmethod
    def parse_expense_pdfs(cls, file_paths):
        """여러 세부사업 PDF를 한 번에 읽고 중복 지급 행을 제거한다."""
        if not file_paths:
            raise ValueError("선택된 지출현황 PDF가 없습니다.")

        all_transactions = []
        sources = []
        messages = []
        for file_path in file_paths:
            txs, metadata, message = cls.parse_expense_pdf(file_path)
            all_transactions.extend(txs)
            sources.append(metadata)
            messages.append(message)

        deduplicated = cls.deduplicate_transactions(all_transactions)
        removed = len(all_transactions) - len(deduplicated)
        summary = (
            f"PDF {len(file_paths)}개에서 지출 {len(deduplicated)}건, "
            f"총 {sum(int(tx.get('amount', 0)) for tx in deduplicated):,}원을 읽었습니다."
        )
        if removed:
            summary += f" 중복 {removed}건은 제외했습니다."
        return deduplicated, sources, summary, messages

    @classmethod
    def _transactions_from_expense_table(
        cls, table, source_file="", page_number=0, department=""
    ) -> List[Dict[str, Any]]:
        if not table or len(table) < 3:
            return []

        header_rows = table[:3]
        header_text = " ".join(
            cls._clean_pdf_text(cell)
            for row in header_rows
            for cell in (row or [])
            if cell
        )
        if "사업명" not in header_text or "통계목" not in header_text or "결의금액" not in header_text:
            return []

        # e호조 지출 집행현황의 22열 표. 헤더명으로 우선 찾고, 병합 헤더는
        # 출력 양식의 고정 열 위치를 사용한다.
        first_header = [cls._clean_pdf_text(v) for v in (table[0] or [])]
        second_header = [cls._clean_pdf_text(v) for v in (table[1] or [])]

        def find_index(name, rows, fallback):
            for row in rows:
                for idx, value in enumerate(row):
                    if name in value:
                        return idx
            return fallback

        business_idx = find_index("사업명", [first_header], 3)
        account_idx = find_index("통계목", [first_header], 4)
        summary_idx = find_index("적요", [first_header], 5)
        proposal_idx = find_index("품의번호", [second_header], 6)
        amount_idx = find_index("결의금액", [second_header], 16)
        payment_order_idx = find_index("지급명령번호", [second_header], 19)
        payment_date_idx = find_index("지급일자", [second_header], 20)
        decision_date_idx = find_index("결의승인일", [second_header], 15)

        start_idx = 2
        results = []
        for row in table[start_idx:]:
            if not row:
                continue
            values = list(row)

            def cell(idx):
                return values[idx] if 0 <= idx < len(values) else None

            business = cls._clean_pdf_text(cell(business_idx))
            account = cls._clean_pdf_text(cell(account_idx))
            summary = cls._clean_pdf_text(cell(summary_idx), preserve_spaces=True)
            amount = cls._clean_amount(cell(amount_idx))
            if not business or not account or amount == 0:
                continue
            if any(term in business for term in ("합계", "소계", "총계")):
                continue

            payment_date = cls._clean_pdf_text(cell(payment_date_idx))
            decision_date = cls._clean_pdf_text(cell(decision_date_idx))
            date_value = payment_date if re.match(r"20\d{2}-\d{2}-\d{2}", payment_date) else decision_date
            payment_order = cls._clean_pdf_text(cell(payment_order_idx))
            proposal_number = cls._clean_pdf_text(cell(proposal_idx))

            results.append({
                "code": payment_order or proposal_number,
                "date": date_value,
                "account": account,
                "sub_account": "미분류",
                "summary": summary,
                "vendor": "",
                "amount": amount,
                "rule_matched": "",
                "detail_project": business,
                "department": department,
                "proposal_number": proposal_number,
                "payment_order_number": payment_order,
                "source_file": source_file,
                "source_page": page_number,
            })
        return results

    @classmethod
    def _clean_pdf_text(cls, value: Any, preserve_spaces: bool = False) -> str:
        if value is None:
            return ""
        text = str(value).replace("\r", "").strip()
        if preserve_spaces:
            text = re.sub(r"\s*\n\s*", "", text)
            return re.sub(r"[ \t]+", " ", text).strip()
        return re.sub(r"\s+", "", text)

    @classmethod
    def transaction_identity(cls, tx: Dict[str, Any]):
        """다른 세부사업 PDF 또는 재업로드에서 같은 지급 행을 식별한다."""
        payment_order = cls._clean_str(tx.get("payment_order_number") or tx.get("code"))
        if payment_order:
            return (
                "payment",
                cls._clean_str(tx.get("department")),
                payment_order,
                cls._clean_str(tx.get("date")),
                int(tx.get("amount", 0)),
            )
        return (
            "content",
            cls._clean_str(tx.get("detail_project")),
            cls._clean_str(tx.get("account")),
            cls._clean_str(tx.get("summary")),
            cls._clean_str(tx.get("date")),
            int(tx.get("amount", 0)),
        )

    @classmethod
    def deduplicate_transactions(cls, transactions):
        seen = set()
        results = []
        for tx in transactions:
            identity = cls.transaction_identity(tx)
            if identity in seen:
                continue
            seen.add(identity)
            results.append(tx)
        return results

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
