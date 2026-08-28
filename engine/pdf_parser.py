"""
추경 세출사업명세서 PDF 파서 (Supplementary Budget PDF Parser)
- 지방자치단체 세출예산 사업명세서 PDF 문서 분석
- 통계목(201-01 등), 기정액, 비교증감, 추경예산액 및 산출기초(○ 세목별 증감) 지능형 파싱
"""

import os
import re
from typing import List, Dict, Any, Tuple, Optional


class SupplementaryPdfParser:
    @classmethod
    def _clean_amount_str(cls, val: str) -> int:
        if not val:
            return 0
        s = re.sub(r'[\,\s원천]', '', val).strip()
        # 음수 처리 (△, ▲, -)
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
        """PDF 파일에서 텍스트를 추출합니다. (pypdf 또는 기본 스트림 파서 사용)"""
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")

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
        except ImportError:
            pass
        except Exception:
            pass

        # 2. PyPDF2 시도
        try:
            import PyPDF2
            reader = PyPDF2.PdfReader(pdf_path)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"
            if extracted_text.strip():
                return extracted_text
        except Exception:
            pass

        # 3. 만약 텍스트 기반 덤프 파일이거나 텍스트 fallback인 경우
        try:
            with open(pdf_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
                # 텍스트 형식인지 확인
                if "세출" in content or "사업명세서" in content or "기정액" in content:
                    return content
        except Exception:
            pass

        if not extracted_text.strip():
            raise ValueError("PDF 파일에서 텍스트를 추출할 수 없습니다. 'pip install pypdf'를 설치하거나 텍스트가 포함된 PDF인지 확인하세요.")

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
        """
        추경 세출사업명세서 PDF 파싱
        반환값:
        {
            "round": 1, # 추경 차수 (1~4회)
            "title": "2026년도 제1회 추가경정예산 세출사업명세서",
            "items": [
                {
                    "account": "201-01 사무관리비",
                    "sub_account": "복사용지 및 사무용품비",
                    "prev_budget": 5000000,
                    "change_amount": 1000000,
                    "revised_budget": 6000000,
                    "reason": "하반기 민원서류 발급 증가에 따른 복사용지 추가 확보"
                },
                ...
            ]
        }
        """
        text = cls.extract_text_from_pdf(pdf_path)
        supp_round = cls.detect_supplementary_round(text)

        items = []

        # 행 단위 분할 파싱
        lines = text.split("\n")
        current_account = ""
        current_prev = 0
        current_change = 0
        current_revised = 0

        # 일반적인 지방자치단체 세출예산 사업명세서 행 매칭 패턴
        # 예: 201-01 사무관리비 | 6,000,000 | 5,000,000 | 1,000,000
        # 산출기초: ○ 복사용지 추가 구입: 5,000,000원 -> 6,000,000원 (+1,000,000원)

        account_pattern = re.compile(r'(\d{3}-\d{2})\s*([^\d\n\r]+)')
        # 산출기초 항목 패턴: '○ 세목명' 또는 '- 세목명' 또는 '□ 세목명'
        sub_item_pattern = re.compile(r'[○\-□•\*]\s*([^:\(\n\r]+)(?:[:\s]+([^\n\r]*))?')

        # 금액 추출 패턴 (천원 단위 또는 원 단위)
        # 예: 1,000,000 또는 1,000천원 또는 (+500,000)
        amount_change_pattern = re.compile(r'([+\-△▲]?\s*[\d,]+)\s*(천원|원)?')

        pending_sub_items = []

        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            # 통계목 헤더 행 감지
            acc_match = account_pattern.search(line_str)
            if acc_match:
                code = acc_match.group(1)
                name = acc_match.group(2).strip()
                current_account = f"{code} {name}".split()[0] + " " + "".join(name.split()[:1])
                
                # 금액들이 같은 줄에 있는 경우 추출
                numbers = re.findall(r'[+\-△▲]?[\d,]{2,}', line_str)
                if len(numbers) >= 3:
                    # [추경예산액, 기정액, 비교증감] 또는 [기정액, 비교증감, 추경예산액]
                    n1 = cls._clean_amount_str(numbers[0])
                    n2 = cls._clean_amount_str(numbers[1])
                    n3 = cls._clean_amount_str(numbers[2])
                    # 단위가 천원인지 원인지 감지 (금액이 100만 이하의 작은 수치면 천원 단위 곱셈)
                    multiplier = 1000 if ("천원" in line_str or max(abs(n1), abs(n2)) < 100000) else 1
                    current_revised = n1 * multiplier
                    current_prev = n2 * multiplier
                    current_change = n3 * multiplier

            # 산출기초 행 감지 (세목별 증감)
            sub_match = sub_item_pattern.search(line_str)
            if sub_match:
                item_title = sub_match.group(1).strip()
                item_detail = sub_match.group(2) or ""
                
                # '합계', '소계' 제외
                if any(x in item_title for x in ["합계", "소계", "재원"]):
                    continue

                # 증감액 추출
                diff_match = re.search(r'\(\s*([+\-△▲]?\s*[\d,]+)\s*(?:천원|원)?\s*\)', line_str)
                diff_amount = 0
                if diff_match:
                    diff_amount = cls._clean_amount_str(diff_match.group(1))
                    if abs(diff_amount) < 100000 and "천원" in line_str:
                        diff_amount *= 1000
                else:
                    # 산식에서 금액 추출
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

        # 만약 산출기초 세부행이 없거나 파싱이 적은 경우 통계목 기준으로 기본 행 생성
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
