"""
세출예산 사업명세서 PDF 파서 (Budget Statement PDF Parser)
- 당해년도 본예산 세출예산명세서 PDF 파싱 (단위사업 - 세부사업 - 편성목 - 통계목 - 산출기초/세목 및 예산액)
- 1~4회 추경 세출사업명세서 PDF 파싱 (기정액, 비교증감, 추경예산액 및 산출기초)
"""

import os
import re
from typing import List, Dict, Any, Tuple, Optional


class BaseBudgetPdfParser:
    """당해년도 본예산 세출예산명세서 PDF 파서"""

    @classmethod
    def _clean_amount_str(cls, val: str) -> int:
        if not val:
            return 0
        s = re.sub(r'[\,\s원천]', '', val).strip()
        is_negative = False
        if s.startswith(('△', '▲', '-')) or ('△' in s) or ('▲' in s):
            is_negative = True
            s = re.sub(r'[△▲\-]', '', s)
        try:
            num = int(float(s))
            return -num if is_negative else num
        except Exception:
            return 0

    @classmethod
    def parse_base_budget_pdf(cls, pdf_path: str) -> Dict[str, Any]:
        """
        본예산 세출예산명세서 PDF 파싱
        반환값:
        {
            "year": 2026,
            "title": "2026년도 본예산 세출예산명세서",
            "items": [
                {
                    "policy_project": "행정운영 및 시정홍보 지원",
                    "unit_project": "기본행정 지원 및 청사 유지관리",
                    "detail_project": "부서 기본운영경비",
                    "category": "200 물건비",
                    "account": "201-01 사무관리비",
                    "sub_account": "복사용지 및 사무용품비",
                    "budget": 5000000,
                    "calculation_basis": "○ 행정업무용 소모품 및 복사용지(A4) 구입: 5,000천원"
                },
                ...
            ]
        }
        """
        text = SupplementaryPdfParser.extract_text_from_pdf(pdf_path)

        # 연도 감지
        year_match = re.search(r'(\d{4})\s*년도?\s*(?:당초|본)?\s*세출예산', text)
        year = int(year_match.group(1)) if year_match else 2026

        items = []

        current_policy = "일반행정 운영"
        current_unit = "기본행정 지원 및 청사관리"
        current_detail = "부서 기본운영경비 지원"
        current_category = "200 물건비"
        current_account = "201-01 사무관리비"
        current_acc_budget = 0

        # 정규식 패턴
        policy_pat = re.compile(r'정책사업\s*[:\s]\s*([^\n\r]+)')
        unit_pat = re.compile(r'단위사업\s*[:\s]\s*([^\n\r]+)')
        detail_pat = re.compile(r'세부사업\s*[:\s]\s*([^\n\r]+)')
        
        # 편성목 패턴 (예: 200 물건비, 300 이전지출)
        category_pat = re.compile(r'(\d{3})\s*([가-힣]+비|[가-힣]+지출)')
        
        # 통계목 패턴 (예: 201-01 사무관리비, 201-02 공공운영비)
        account_pat = re.compile(r'(\d{3}-\d{2})\s*([^\d\n\r]+)')
        
        # 산출기초 세목 패턴 (예: ○ 복사용지 및 토너 구입: 5,000천원)
        sub_item_pat = re.compile(r'[○\-□•\*]\s*([^:\(\n\r]+)(?:[:\s]+([^\n\r]*))?')

        lines = text.split("\n")

        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            # 1. 정책사업 감지
            p_m = policy_pat.search(line_str)
            if p_m:
                current_policy = p_m.group(1).strip()
                continue

            # 2. 단위사업 감지
            u_m = unit_pat.search(line_str)
            if u_m:
                current_unit = u_m.group(1).strip()
                continue

            # 3. 세부사업 감지
            d_m = detail_pat.search(line_str)
            if d_m:
                current_detail = d_m.group(1).strip()
                continue

            # 4. 통계목 감지
            acc_m = account_pat.search(line_str)
            if acc_m:
                code = acc_m.group(1)
                name = acc_m.group(2).strip().split()[0]
                current_account = f"{code} {name}"
                
                # 편성목 추론 (앞 3자리 기준)
                c_code = code.split("-")[0]
                cat_map = {
                    "201": "200 물건비(일반운영비)",
                    "202": "200 물건비(여비)",
                    "203": "200 물건비(업무추진비)",
                    "206": "200 물건비(재료비)",
                    "307": "300 이전지출(민간이전)",
                    "401": "400 시설비및부대비",
                    "405": "400 자산취득비"
                }
                current_category = cat_map.get(c_code, f"{c_code[:1]}00 예산목")

                # 통계목 예산액 추출
                nums = re.findall(r'[\d,]{3,}', line_str)
                if nums:
                    raw_b = cls._clean_amount_str(nums[0])
                    mult = 1000 if ("천원" in line_str or raw_b < 100000) else 1
                    current_acc_budget = raw_b * mult
                continue

            # 5. 산출기초(세목) 감지
            sub_m = sub_item_pat.search(line_str)
            if sub_m:
                sub_title = sub_m.group(1).strip()
                sub_calc = sub_m.group(2) or ""

                if any(x in sub_title for x in ["합계", "소계", "재원", "국비", "시비", "구비"]):
                    continue

                # 산출기초 내 금액 추출
                amt = 0
                nums = re.findall(r'[\d,]{3,}', line_str)
                if nums:
                    raw_a = cls._clean_amount_str(nums[-1])
                    mult = 1000 if ("천원" in line_str or raw_a < 100000) else 1
                    amt = raw_a * mult

                items.append({
                    "policy_project": current_policy,
                    "unit_project": current_unit,
                    "detail_project": current_detail,
                    "category": current_category,
                    "account": current_account,
                    "sub_account": sub_title,
                    "budget": amt,
                    "calculation_basis": line_str
                })

        # 만약 산출기초가 없고 통계목만 추출된 경우 기본값 처리
        if not items and current_account:
            items.append({
                "policy_project": current_policy,
                "unit_project": current_unit,
                "detail_project": current_detail,
                "category": current_category,
                "account": current_account,
                "sub_account": "기본 운영비",
                "budget": current_acc_budget,
                "calculation_basis": f"○ {current_account} 기본예산 편성"
            })

        return {
            "year": year,
            "title": f"{year}년도 본예산 세출예산명세서",
            "items": items,
            "total_budget": sum(it["budget"] for it in items)
        }


class SupplementaryPdfParser:
    @classmethod
    def _clean_amount_str(cls, val: str) -> int:
        if not val:
            return 0
        s = re.sub(r'[\,\s원천]', '', val).strip()
        is_negative = False
        if s.startswith(('△', '▲', '-')) or ('△' in s) or ('▲' in s):
            is_negative = True
            s = re.sub(r'[△▲\-]', '', s)
        try:
            num = int(float(s))
            return -num if is_negative else num
        except Exception:
            return 0

    @classmethod
    def extract_text_from_pdf(cls, pdf_path: str) -> str:
        """PDF 파일에서 텍스트를 추출합니다. (pypdf 또는 텍스트 fallback)"""
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {pdf_path}")

        extracted_text = ""

        # 1. pypdf 시도
        try:
            import pypdf
            reader = pypdf.PdfReader(pdf_path)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"
            if extracted_text.strip():
                return extracted_text
        except Exception:
            pass

        # 2. 텍스트 파일인 경우 직접 읽기
        try:
            with open(pdf_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
                if "세출" in content or "사업명세서" in content or "예산" in content:
                    return content
        except Exception:
            pass

        if not extracted_text.strip():
            raise ValueError("PDF/TXT 파일에서 유효한 텍스트를 추출할 수 없습니다.")

        return extracted_text

    @classmethod
    def detect_supplementary_round(cls, text: str) -> int:
        """문서 텍스트에서 몇 회 추경인지 자동 감지 (1~4회)"""
        match = re.search(r'제\s*([1-4일이삼사])\s*회\s*(추가경정|추경)', text)
        if match:
            r_str = match.group(1)
            mapping = {"1": 1, "일": 1, "2": 2, "이": 2, "3": 3, "삼": 3, "4": 4, "사": 4}
            return mapping.get(r_str, 1)
        return 1

    @classmethod
    def parse_supplementary_pdf(cls, pdf_path: str) -> Dict[str, Any]:
        """추경 세출사업명세서 PDF 파싱"""
        text = cls.extract_text_from_pdf(pdf_path)
        supp_round = cls.detect_supplementary_round(text)

        items = []
        lines = text.split("\n")
        current_account = ""
        current_prev = 0
        current_change = 0
        current_revised = 0

        account_pattern = re.compile(r'(\d{3}-\d{2})\s*([^\d\n\r]+)')
        sub_item_pattern = re.compile(r'[○\-□•\*]\s*([^:\(\n\r]+)(?:[:\s]+([^\n\r]*))?')

        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            acc_match = account_pattern.search(line_str)
            if acc_match:
                code = acc_match.group(1)
                name = acc_match.group(2).strip().split()[0]
                current_account = f"{code} {name}"

                numbers = re.findall(r'[+\-△▲]?[\d,]{2,}', line_str)
                if len(numbers) >= 3:
                    n1 = cls._clean_amount_str(numbers[0])
                    n2 = cls._clean_amount_str(numbers[1])
                    n3 = cls._clean_amount_str(numbers[2])
                    multiplier = 1000 if ("천원" in line_str or max(abs(n1), abs(n2)) < 100000) else 1
                    current_revised = n1 * multiplier
                    current_prev = n2 * multiplier
                    current_change = n3 * multiplier

            sub_match = sub_item_pattern.search(line_str)
            if sub_match:
                item_title = sub_match.group(1).strip()
                item_detail = sub_match.group(2) or ""

                if any(x in item_title for x in ["합계", "소계", "재원"]):
                    continue

                diff_match = re.search(r'\(\s*([+\-△▲]?\s*[\d,]+)\s*(?:천원|원)?\s*\)', line_str)
                diff_amount = 0
                if diff_match:
                    diff_amount = cls._clean_amount_str(diff_match.group(1))
                    if abs(diff_amount) < 100000 and "천원" in line_str:
                        diff_amount *= 1000
                else:
                    nums = re.findall(r'[+\-△▲]?[\d,]{3,}', line_str)
                    if nums:
                        diff_amount = cls._clean_amount_str(nums[-1])

                items.append({
                    "account": current_account or "201-01 사무관리비",
                    "sub_account": item_title,
                    "prev_budget": 0,
                    "change_amount": diff_amount,
                    "revised_budget": 0,
                    "reason": item_detail.strip() or f"제{supp_round}회 추경 예산 조정"
                })

        if not items and current_account:
            items.append({
                "account": current_account,
                "sub_account": "기타 세부사업",
                "prev_budget": current_prev,
                "change_amount": current_change,
                "revised_budget": current_revised,
                "reason": f"제{supp_round}회 추경 예산 조정"
            })

        return {
            "round": supp_round,
            "title": f"2026년도 제{supp_round}회 추가경정예산 세출사업명세서",
            "items": items,
            "raw_text_length": len(text)
        }
