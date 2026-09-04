import unittest
from unittest.mock import patch

try:
    from main import BudgetApp, SupplementaryReviewDialog
except ImportError:
    BudgetApp = None
    SupplementaryReviewDialog = None


@unittest.skipIf(SupplementaryReviewDialog is None, "Tkinter runtime is unavailable")
class SupplementaryWorkflowTests(unittest.TestCase):
    @staticmethod
    def _base_items():
        budget = {
            "policy_project": "행정운영경비(차량등록사업소)",
            "unit_project": "인력운영비",
            "detail_project": "인력운영비",
            "detail_project_budget": 5_386_285_000,
            "category": "101 인건비",
            "category_budget": 5_332_544_000,
            "account": "101-01 보수",
            "sub_account": "통계목 전체",
            "budget": 4_564_006_000,
            "budget_level": "account",
        }
        supplementary = {
            **budget,
            "base_budget": budget["budget"],
            "supplements": {"1": 0, "2": 0, "3": 0, "4": 0},
            "reasons": {},
            "final_budget": budget["budget"],
        }
        return budget, supplementary

    def test_applies_changed_statistic_to_existing_base_budget(self):
        budget, supplementary = self._base_items()
        incoming = {
            "policy_project": budget["policy_project"],
            "unit_project": budget["unit_project"],
            "detail_project": budget["detail_project"],
            "category": budget["category"],
            "account": budget["account"],
            "sub_account": "통계목 전체",
            "prev_budget": 4_564_006_000,
            "change_amount": 90_240_000,
            "revised_budget": 4_654_246_000,
            "budget_level": "account",
            "reason": "제3회 추경",
        }

        dialog = object.__new__(SupplementaryReviewDialog)
        dialog.combo_round = type("Combo", (), {"current": lambda self: 2})()
        dialog.parsed_data = {"items": [incoming]}
        dialog.current_budget_master = [budget]
        dialog.current_supp_items = [supplementary]
        dialog.result = None
        dialog.destroy = lambda: None
        dialog._on_apply()

        self.assertEqual(dialog.result["round"], 3)
        self.assertEqual(dialog.result["updated_budget_master"][0]["budget"], 4_654_246_000)
        updated_supp = dialog.result["updated_supp_items"][0]
        self.assertEqual(updated_supp["supplements"]["3"], 90_240_000)
        self.assertEqual(updated_supp["final_budget"], 4_654_246_000)

    def test_blocks_supplementary_upload_without_confirmed_base_budget(self):
        app = object.__new__(BudgetApp)
        app.base_budget_confirmed = False
        app.budget_master = []
        app.supplementary_budgets = []
        app.log = lambda *args, **kwargs: None

        with patch("main.messagebox.showwarning") as warning, patch(
            "main.filedialog.askopenfilename"
        ) as file_dialog:
            app._on_upload_supplementary_pdf()

        warning.assert_called_once()
        file_dialog.assert_not_called()


if __name__ == "__main__":
    unittest.main()
