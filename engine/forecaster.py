"""
12.31 연말 지출 및 불용액 예측 엔진 (Year-End Expenditure & Balance Forecaster)
- 단위사업 - 세부사업 - 편성목 - 통계목 - 세목 계층별 집계
- 정기 고정비(월정액 * 잔여월수) 및 하반기 예정액 계산
- 12.31 기준 예상 지출총액 및 예상 잔액(불용액) 시뮬레이션
- 1월 ~ 12월 월별 지출 예산 통계 매트릭스 집계
- 1~4회 추경(추가경정예산) 변동 현황 및 본예산 비교 집계
"""

from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple


class Forecaster:
    @staticmethod
    def _detail_key(value: Any) -> str:
        return "".join(str(value or "").split())

    @classmethod
    def _budget_identity(cls, item: Dict[str, Any]) -> Tuple[str, str, str]:
        return (
            cls._detail_key(item.get("detail_project")),
            item.get("account", "기타"),
            item.get("sub_account", "미분류"),
        )

    @classmethod
    def _build_spending_maps(cls, transactions: List[Dict[str, Any]]):
        exact = {}
        legacy = {}
        for tx in transactions:
            account = tx.get("account", "기타")
            sub_account = tx.get("sub_account", "미분류")
            amount = int(tx.get("amount", 0))
            detail = cls._detail_key(tx.get("detail_project"))
            if detail:
                key = (detail, account, sub_account)
                exact[key] = exact.get(key, 0) + amount
            else:
                key = (account, sub_account)
                legacy[key] = legacy.get(key, 0) + amount
        return exact, legacy

    @classmethod
    def _legacy_budget_counts(cls, budget_items: List[Dict[str, Any]]):
        counts = {}
        for item in budget_items:
            key = (item.get("account", "기타"), item.get("sub_account", "미분류"))
            counts[key] = counts.get(key, 0) + 1
        return counts

    @staticmethod
    def calculate_remaining_months(base_date: Optional[str] = None) -> int:
        if base_date:
            try:
                clean_date = base_date.replace(".", "-").replace("/", "-")
                dt = datetime.strptime(clean_date.strip()[:10], "%Y-%m-%d")
            except Exception:
                dt = datetime.now()
        else:
            dt = datetime.now()

        current_month = dt.month
        remaining = 12 - current_month
        return max(0, remaining)

    @classmethod
    def calculate_monthly_matrix(
        cls,
        budget_items: List[Dict[str, Any]],
        transactions: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        budget_keys = {cls._budget_identity(item) for item in budget_items}
        legacy_counts = cls._legacy_budget_counts(budget_items)
        monthly_map = {}
        legacy_monthly_map = {}
        unbudgeted_map = {}
        for tx in transactions:
            acc = tx.get("account", "기타")
            sub_acc = tx.get("sub_account", "미분류")
            detail = cls._detail_key(tx.get("detail_project"))
            exact_key = (detail, acc, sub_acc)
            legacy_key = (acc, sub_acc)

            date_str = str(tx.get("date", "")).replace(".", "-").replace("/", "-")
            month = 0
            try:
                parts = date_str.split("-")
                if len(parts) >= 2:
                    month = int(parts[1])
            except Exception:
                month = 0

            if 1 <= month <= 12:
                amount = int(tx.get("amount", 0))
                if detail:
                    target = monthly_map.setdefault(exact_key, [0] * 12)
                    target[month - 1] += amount
                    if exact_key not in budget_keys:
                        unmatched = unbudgeted_map.setdefault(exact_key, [0] * 12)
                        unmatched[month - 1] += amount
                else:
                    target = legacy_monthly_map.setdefault(legacy_key, [0] * 12)
                    target[month - 1] += amount
                    if legacy_counts.get(legacy_key, 0) != 1:
                        unmatched = unbudgeted_map.setdefault(("", acc, sub_acc), [0] * 12)
                        unmatched[month - 1] += amount

        matrix_rows = []
        monthly_totals = [0] * 12
        total_budget = 0
        total_spent = 0

        for b in budget_items:
            acc = b.get("account", "기타")
            sub_acc = b.get("sub_account", "기타")
            budget = int(b.get("budget", 0))
            exact_key = cls._budget_identity(b)
            legacy_key = (acc, sub_acc)
            m_values = list(monthly_map.get(exact_key, [0] * 12))
            if legacy_counts.get(legacy_key, 0) == 1:
                legacy_values = legacy_monthly_map.get(legacy_key, [0] * 12)
                m_values = [m_values[i] + legacy_values[i] for i in range(12)]
            row_sum = sum(m_values)
            balance = budget - row_sum
            exec_rate = round((row_sum / budget * 100), 1) if budget > 0 else 0.0

            for i in range(12):
                monthly_totals[i] += m_values[i]
            total_budget += budget
            total_spent += row_sum

            matrix_rows.append({
                "unit_project": b.get("unit_project", "기본행정 지원"),
                "detail_project": b.get("detail_project", "부서 기본운영경비"),
                "category": b.get("category", "물건비"),
                "account": acc,
                "sub_account": sub_acc,
                "budget": budget,
                "months": m_values,
                "total_spent": row_sum,
                "balance": balance,
                "exec_rate": exec_rate
            })

        for (detail, acc, sub_acc), m_values in unbudgeted_map.items():
            row_sum = sum(m_values)
            for i in range(12):
                monthly_totals[i] += m_values[i]
            total_spent += row_sum
            matrix_rows.append({
                "unit_project": "-",
                "detail_project": detail or "-",
                "category": "-",
                "account": acc,
                "sub_account": sub_acc,
                "budget": 0,
                "months": m_values,
                "total_spent": row_sum,
                "balance": -row_sum,
                "exec_rate": 0.0
            })

        overall_exec_rate = round((total_spent / total_budget * 100), 1) if total_budget > 0 else 0.0

        return {
            "rows": matrix_rows,
            "monthly_totals": monthly_totals,
            "total_budget": total_budget,
            "total_spent": total_spent,
            "total_balance": total_budget - total_spent,
            "overall_exec_rate": overall_exec_rate
        }

    @classmethod
    def calculate_supplementary_matrix(
        cls,
        supplementary_items: List[Dict[str, Any]],
        exact_spent: Dict[Tuple[str, str, str], int],
        legacy_spent: Optional[Dict[Tuple[str, str], int]] = None,
    ) -> Dict[str, Any]:
        legacy_spent = legacy_spent or {}
        legacy_counts = cls._legacy_budget_counts(supplementary_items)
        rows = []
        tot_base = 0
        tot_r1 = 0
        tot_r2 = 0
        tot_r3 = 0
        tot_r4 = 0
        tot_final = 0
        tot_spent = 0

        for item in supplementary_items:
            acc = item.get("account", "")
            sub_acc = item.get("sub_account", "")
            base_b = int(item.get("base_budget", 0))
            supps = item.get("supplements", {})
            r1 = int(supps.get("1", 0))
            r2 = int(supps.get("2", 0))
            r3 = int(supps.get("3", 0))
            r4 = int(supps.get("4", 0))

            final_b = base_b + r1 + r2 + r3 + r4
            exact_key = cls._budget_identity(item)
            legacy_key = (acc, sub_acc)
            spent = exact_spent.get(exact_key, 0)
            if legacy_counts.get(legacy_key, 0) == 1:
                spent += legacy_spent.get(legacy_key, 0)
            balance = final_b - spent
            exec_rate = round((spent / final_b * 100), 1) if final_b > 0 else 0.0

            reasons_dict = item.get("reasons", {})
            reasons_str = " / ".join(f"[{k}회추경] {v}" for k, v in reasons_dict.items() if v)

            tot_base += base_b
            tot_r1 += r1
            tot_r2 += r2
            tot_r3 += r3
            tot_r4 += r4
            tot_final += final_b
            tot_spent += spent

            rows.append({
                "unit_project": item.get("unit_project", "기본행정 지원"),
                "detail_project": item.get("detail_project", "부서 기본운영경비"),
                "account": acc,
                "sub_account": sub_acc,
                "base_budget": base_b,
                "r1": r1,
                "r2": r2,
                "r3": r3,
                "r4": r4,
                "final_budget": final_b,
                "spent": spent,
                "balance": balance,
                "exec_rate": exec_rate,
                "reason": reasons_str
            })

        return {
            "rows": rows,
            "total_base": tot_base,
            "total_r1": tot_r1,
            "total_r2": tot_r2,
            "total_r3": tot_r3,
            "total_r4": tot_r4,
            "total_final": tot_final,
            "total_spent": tot_spent,
            "total_balance": tot_final - tot_spent,
            "overall_exec_rate": round((tot_spent / tot_final * 100), 1) if tot_final > 0 else 0.0
        }

    @classmethod
    def simulate(
        cls,
        budget_items: List[Dict[str, Any]],
        transactions: List[Dict[str, Any]],
        recurring_plans: Optional[List[Dict[str, Any]]] = None,
        scheduled_plans: Optional[List[Dict[str, Any]]] = None,
        supplementary_items: Optional[List[Dict[str, Any]]] = None,
        base_date: Optional[str] = None
    ) -> Dict[str, Any]:
        remaining_months = cls.calculate_remaining_months(base_date)

        exact_spent, legacy_spent = cls._build_spending_maps(transactions)
        legacy_counts = cls._legacy_budget_counts(budget_items)

        recurring_map = {}
        if recurring_plans:
            for r in recurring_plans:
                key = (r.get("account"), r.get("sub_account"))
                monthly = int(r.get("monthly_amount", 0))
                recurring_map[key] = recurring_map.get(key, 0) + monthly

        scheduled_map = {}
        if scheduled_plans:
            for s in scheduled_plans:
                key = (s.get("account"), s.get("sub_account"))
                amt = int(s.get("amount", 0))
                scheduled_map[key] = scheduled_map.get(key, 0) + amt

        summary_items = []
        total_budget = 0
        total_spent = 0
        total_forecast_spent = 0
        total_forecast_balance = 0

        detail_groups = {}

        for b in budget_items:
            unit_p = b.get("unit_project", "기본행정 지원")
            det_p = b.get("detail_project", "부서 기본운영경비")
            cat = b.get("category", "물건비")
            acc = b.get("account", "기타")
            sub_acc = b.get("sub_account", "기타")
            budget = int(b.get("budget", 0))
            exact_key = cls._budget_identity(b)
            legacy_key = (acc, sub_acc)
            plan_key = (acc, sub_acc)

            actual_spent = exact_spent.get(exact_key, 0)
            if legacy_counts.get(legacy_key, 0) == 1:
                actual_spent += legacy_spent.get(legacy_key, 0)
            current_balance = budget - actual_spent
            exec_rate = round((actual_spent / budget * 100), 1) if budget > 0 else 0.0

            monthly_rec = recurring_map.get(plan_key, 0)
            rec_spent = monthly_rec * remaining_months
            sched_spent = scheduled_map.get(plan_key, 0)

            forecast_total_spent = actual_spent + rec_spent + sched_spent
            forecast_balance = budget - forecast_total_spent
            forecast_exec_rate = round((forecast_total_spent / budget * 100), 1) if budget > 0 else 0.0

            if forecast_balance < 0:
                status = "초과위험(부족)"
                status_badge = "danger"
            elif forecast_exec_rate < 85.0 and budget >= 1000000:
                status = "불용위험(과다)"
                status_badge = "warning"
            else:
                status = "양호"
                status_badge = "success"

            item_data = {
                "unit_project": unit_p,
                "detail_project": det_p,
                "category": cat,
                "account": acc,
                "sub_account": sub_acc,
                "budget": budget,
                "actual_spent": actual_spent,
                "current_balance": current_balance,
                "exec_rate": exec_rate,
                "monthly_recurring": monthly_rec,
                "remaining_recurring": rec_spent,
                "scheduled_spent": sched_spent,
                "forecast_total_spent": forecast_total_spent,
                "forecast_balance": forecast_balance,
                "forecast_exec_rate": forecast_exec_rate,
                "status": status,
                "status_badge": status_badge,
                "note": b.get("note", "")
            }

            summary_items.append(item_data)

            if det_p not in detail_groups:
                detail_groups[det_p] = {
                    "detail_project": det_p,
                    "budget": 0,
                    "actual_spent": 0,
                    "forecast_total_spent": 0,
                    "forecast_balance": 0
                }
            detail_groups[det_p]["budget"] += budget
            detail_groups[det_p]["actual_spent"] += actual_spent
            detail_groups[det_p]["forecast_total_spent"] += forecast_total_spent
            detail_groups[det_p]["forecast_balance"] += forecast_balance

            total_budget += budget
            total_spent += actual_spent
            total_forecast_spent += forecast_total_spent
            total_forecast_balance += forecast_balance

        budget_keys = {cls._budget_identity(item) for item in budget_items}
        unclassified_spent = sum(
            amount for key, amount in exact_spent.items() if key not in budget_keys
        )
        unclassified_spent += sum(
            amount
            for key, amount in legacy_spent.items()
            if legacy_counts.get(key, 0) != 1
        )

        overall_exec_rate = round((total_spent / total_budget * 100), 1) if total_budget > 0 else 0.0
        overall_forecast_exec_rate = round((total_forecast_spent / total_budget * 100), 1) if total_budget > 0 else 0.0

        monthly_matrix = cls.calculate_monthly_matrix(budget_items, transactions)
        supp_matrix = {}
        if supplementary_items:
            supp_matrix = cls.calculate_supplementary_matrix(
                supplementary_items, exact_spent, legacy_spent
            )

        return {
            "remaining_months": remaining_months,
            "items": summary_items,
            "detail_groups": list(detail_groups.values()),
            "total_budget": total_budget,
            "total_spent": total_spent,
            "current_balance": total_budget - total_spent,
            "overall_exec_rate": overall_exec_rate,
            "total_forecast_spent": total_forecast_spent,
            "total_forecast_balance": total_forecast_balance,
            "overall_forecast_exec_rate": overall_forecast_exec_rate,
            "unclassified_spent": unclassified_spent,
            "monthly_matrix": monthly_matrix,
            "supplementary_matrix": supp_matrix
        }
