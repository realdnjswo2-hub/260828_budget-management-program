import unittest

from engine.pdf_parser import BaseBudgetPdfParser, SupplementaryPdfParser


class BaseBudgetPdfParserTests(unittest.TestCase):
    def test_hierarchical_report_extracts_budget_levels_in_won(self):
        text = """
2026년도 본예산 일반회계
부서: 차량등록사업소
정책: 차량등록사업소 운영
단위: 교통민원 주민 접근성 제고 (단위:천원)
차량등록사업소 100,000 90,000 10,000
차량등록사업소 운영 100,000 90,000 10,000
교통민원 주민 접근성 제고 60,000 55,000 5,000
현장민원처리센터 활성화 60,000 55,000 5,000
201 일반운영비 60,000 55,000 5,000
01 사무관리비 20,000 18,000 2,000
02 공공운영비 40,000 37,000 3,000
교통민원 서비스 향상 40,000 35,000 5,000
쾌적한 민원환경 조성 40,000 35,000 5,000
203 업무추진비 40,000 35,000 5,000
04 부서운영업무추진비 40,000 35,000 5,000
"""

        parsed = BaseBudgetPdfParser._parse_hierarchical_report(text, 2026)

        self.assertEqual(parsed["import_level"], "통계목")
        self.assertEqual(parsed["total_budget"], 100_000_000)
        self.assertEqual(len(parsed["items"]), 3)

        first = parsed["items"][0]
        self.assertEqual(first["detail_project"], "현장민원처리센터 활성화")
        self.assertEqual(first["detail_project_budget"], 60_000_000)
        self.assertEqual(first["category"], "201 일반운영비")
        self.assertEqual(first["category_budget"], 60_000_000)
        self.assertEqual(first["account"], "201-01 사무관리비")
        self.assertEqual(first["budget"], 20_000_000)
        self.assertEqual(first["budget_level"], "account")

        last = parsed["items"][-1]
        self.assertEqual(last["unit_project"], "교통민원 서비스 향상")
        self.assertEqual(last["detail_project"], "쾌적한 민원환경 조성")
        self.assertEqual(last["account"], "203-04 부서운영업무추진비")

    def test_hierarchical_supplementary_extracts_only_changed_statistic(self):
        text = """
2026년도 추경 3 회 일반회계
부서: 차량등록사업소
정책: 행정운영경비(차량등록사업소)
단위: 인력운영비 (단위:천원)
차량등록사업소 6,670,552 6,580,312 90,240
행정운영경비(차량등록사업소) 5,579,065 5,488,825 90,240
인력운영비 5,476,525 5,386,285 90,240
인력운영비 5,476,525 5,386,285 90,240
101 인건비 5,422,784 5,332,544 90,240
01 보수 4,654,246 4,564,006 90,240
02 기타직보수 649,861 649,861 0
"""

        parsed = SupplementaryPdfParser._parse_hierarchical_supplementary(text, 2026, 3)

        self.assertEqual(parsed["round"], 3)
        self.assertEqual(parsed["total_change"], 90_240_000)
        self.assertEqual(len(parsed["items"]), 1)
        item = parsed["items"][0]
        self.assertEqual(item["detail_project"], "인력운영비")
        self.assertEqual(item["category"], "101 인건비")
        self.assertEqual(item["account"], "101-01 보수")
        self.assertEqual(item["prev_budget"], 4_564_006_000)
        self.assertEqual(item["change_amount"], 90_240_000)
        self.assertEqual(item["revised_budget"], 4_654_246_000)

    def test_detects_both_supplementary_round_title_styles(self):
        self.assertEqual(SupplementaryPdfParser.detect_supplementary_round("제3회 추경"), 3)
        self.assertEqual(SupplementaryPdfParser.detect_supplementary_round("추경 3 회"), 3)


if __name__ == "__main__":
    unittest.main()
