import os
import tempfile
import unittest

from engine.forecaster import Forecaster
from engine.parser import EHojoParser
from engine.project_store import ProjectStore


class ExpensePdfParserTests(unittest.TestCase):
    @staticmethod
    def _table():
        return [
            [
                "회계구분", "부서명", "경비구분", "사업명", "통계목", "적요",
                "품의정보", None, None, "원인정보", None, None, None, None, None,
                "결의정보", None, None, None, "지급정보", None, "품의유형",
            ],
            [
                None, None, None, None, None, None, "품의\n번호", "품의금액", "발의일자",
                "원인번호", "승인일자", "원인금액", "원인상태", "원인요청일", "원인구분",
                "결의승인일", "결의금액", "결의상태", "결의구분", "지급\n명령번호", "지급일자", None,
            ],
            [
                "일반회계", "차량등록사업소", "일반지출", "현장민원처리센터 활성화",
                "사무관리비", "폐기물\n수거 용역비", "4", "140,000", "2026-01-02",
                "00000004", "2026-01-02", "140,000", "확정", "2026-01-02", "일반지출",
                "2026-01-05", "140,000", "결의확정", "일반지출", "00000002", "2026-01-05", "일반",
            ],
        ]

    def test_reads_required_columns_from_expense_table(self):
        rows = EHojoParser._transactions_from_expense_table(
            self._table(), source_file="sample.pdf", page_number=1, department="차량등록사업소"
        )

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["detail_project"], "현장민원처리센터활성화")
        self.assertEqual(rows[0]["account"], "사무관리비")
        self.assertEqual(rows[0]["date"], "2026-01-05")
        self.assertEqual(rows[0]["amount"], 140_000)
        self.assertEqual(rows[0]["payment_order_number"], "00000002")

    def test_deduplicates_same_payment_order(self):
        rows = EHojoParser._transactions_from_expense_table(self._table())
        duplicate = dict(rows[0])
        self.assertEqual(len(EHojoParser.deduplicate_transactions(rows + [duplicate])), 1)


class PortableProjectTests(unittest.TestCase):
    def test_round_trip_preserves_budget_and_transactions(self):
        state = {
            "base_budget_confirmed": True,
            "base_budget_year": 2026,
            "budget_master": [{"account": "201-01 사무관리비", "budget": 1_000_000}],
            "supplementary_budgets": [],
            "rules": [],
            "recurring_plans": [],
            "scheduled_plans": [],
            "raw_transactions": [{"code": "1", "amount": 140_000}],
            "transactions": [{"code": "1", "amount": 140_000}],
            "expense_sources": [{"file_name": "현장민원.pdf"}],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            path = ProjectStore.save(os.path.join(temp_dir, "2026_budget"), state)
            loaded = ProjectStore.load(path)

        self.assertTrue(path.endswith(".ebudget"))
        self.assertEqual(loaded["base_budget_year"], 2026)
        self.assertEqual(loaded["raw_transactions"][0]["amount"], 140_000)
        self.assertEqual(loaded["expense_sources"][0]["file_name"], "현장민원.pdf")


class DetailProjectAggregationTests(unittest.TestCase):
    def test_same_account_in_different_projects_is_not_double_counted(self):
        budgets = [
            {
                "detail_project": "사업 A", "account": "201-01 사무관리비",
                "sub_account": "통계목 전체", "budget": 1_000_000,
            },
            {
                "detail_project": "사업 B", "account": "201-01 사무관리비",
                "sub_account": "통계목 전체", "budget": 2_000_000,
            },
        ]
        transactions = [
            {
                "detail_project": "사업 A", "account": "201-01 사무관리비",
                "sub_account": "통계목 전체", "amount": 100_000, "date": "2026-01-10",
            },
            {
                "detail_project": "사업 B", "account": "201-01 사무관리비",
                "sub_account": "통계목 전체", "amount": 250_000, "date": "2026-02-10",
            },
        ]

        result = Forecaster.simulate(budgets, transactions)

        self.assertEqual([row["actual_spent"] for row in result["items"]], [100_000, 250_000])
        self.assertEqual(result["total_spent"], 350_000)
        self.assertEqual(result["monthly_matrix"]["monthly_totals"][:2], [100_000, 250_000])


if __name__ == "__main__":
    unittest.main()
