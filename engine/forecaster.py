"""
12.31 연말 지출 및 불용액 예측 엔진 (Year-End Expenditure & Balance Forecaster)
- 정기 고정비(월정액 * 잔여월수) 자동 계산
- 하반기 집행예정 사업비 반영
- 12.31 기준 예상 지출총액 및 예상 잔액(불용액) 시뮬레이션
- 1월 ~ 12월 월별 지출 예산 통계 매트릭스 집계
- 1~4회 추경(추가경정예산) 변동 현황 및 본예산 비교 집계
"""

from datetime import datetime
from typing import List, Dict, Any, Optional


class Forecaster:
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
        monthly_map = {}
        for tx in transactions:
            acc = tx.get("account", "기타")
            sub_acc = tx.get("sub_account", "미분류")
            key = (acc, sub_acc)
            if key not in monthly_map:
                monthly_map[key] = [0] * 12

            date_str = str(tx.get("date", "")).replace(".", "-").replace("/", "-")
            month = 0
            try:
                parts = date_str.split("-")
                if len(parts) >= 2:
                    month = int(parts[1])
            except Exception:
                month = 0

            if 1 <= month <= 12:
                monthly_map[key][month - 1] += int(tx.get("amount", 0))

        matrix_rows = []
        monthly_totals = [0] * 12
        total_budget = 0
        total_spent = 0

        for b in budget_items:
            acc = b.get("account", "기타")
            sub_acc = b.get("sub_account", "기타")
            budget = int(b.get("budget", 0))
            key = (acc, sub_acc)

            m_values = monthly_map.get(key, [0] * 12)
            row_sum = sum(m_values)
            balance = budget - row_sum
            exec_rate = round((row_sum / budget * 100), 1) if budget > 0 else 0.0

            for i in range(12):
                monthly_totals[i] += m_values[i]
            total_budget += budget
            total_spent += row_sum

            matrix_rows.append({
                "account": acc,
                "sub_account": sub_acc,
                "budget": budget,
                "months": m_values,
                "total_spent": row_sum,
                "balance": balance,
                "exec_rate": exec_rate
            })

        for (acc, sub_acc), m_values in monthly_map.items():
            if not any(b.get("account") == acc and b.get("sub_account") == sub_acc for b in budget_items):
                row_sum = sum(m_values)
                for i in range(12):
                    monthly_totals[i] += m_values[i]
                total_spent += row_sum
                matrix_rows.append({
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
        spent_map: Dict[Tuple[str, str], int]
    ) -> Dict[str, Any]:
        """1~4회 추경 변동 현황표 집계"""
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
            spent = spent_map.get((acc, sub_acc), 0)
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

        # 1. 세목별 기집행액 집계
        spent_map = {}
        for tx in transactions:
            acc = tx.get("account", "기타")
            sub_acc = tx.get("sub_account", "미분류")
            key = (acc, sub_acc)
            spent_map[key] = spent_map.get(key, 0) + int(tx.get("amount", 0))

        # 2. 세목별 정기지출(월정액) 집계
        recurring_map = {}
        if recurring_plans:
            for r in recurring_plans:
                key = (r.get("account"), r.get("sub_account"))
                monthly = int(r.get("monthly_amount", 0))
                recurring_map[key] = recurring_map.get(key, 0) + monthly

        # 3. 세목별 집행예정 사업비 집계
        scheduled_map = {}
        if scheduled_plans:
            for s in scheduled_plans:
                key = (s.get("account"), s.get("sub_account"))
                amt = int(s.get("amount", 0))
                scheduled_map[key] = scheduled_map.get(key, 0) + amt

        # 4. 세목별 시뮬레이션 결과 생성
        summary_items = []
        total_budget = 0
        total_spent = 0
        total_forecast_spent = 0
        total_forecast_balance = 0

        account_groups = {}

        for b in budget_items:
            acc = b.get("account", "기타")
            sub_acc = b.get("sub_account", "기타")
            budget = int(b.get("budget", 0))
            key = (acc, sub_acc)

            actual_spent = spent_map.get(key, 0)
            current_balance = budget - actual_spent
            exec_rate = round((actual_spent / budget * 100), 1) if budget > 0 else 0.0

            monthly_rec = recurring_map.get(key, 0)
            rec_spent = monthly_rec * remaining_months
            sched_spent = scheduled_map.get(key, 0)

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

            if acc not in account_groups:
                account_groups[acc] = {
                    "account": acc,
                    "budget": 0,
                    "actual_spent": 0,
                    "forecast_total_spent": 0,
                    "forecast_balance": 0
                }
            account_groups[acc]["budget"] += budget
            account_groups[acc]["actual_spent"] += actual_spent
            account_groups[acc]["forecast_total_spent"] += forecast_total_spent
            account_groups[acc]["forecast_balance"] += forecast_balance

            total_budget += budget
            total_spent += actual_spent
            total_forecast_spent += forecast_total_spent
            total_forecast_balance += forecast_balance

        unclassified_spent = 0
        for (acc, sub_acc), amt in spent_map.items():
            if not any(b.get("account") == acc and b.get("sub_account") == sub_acc for b in budget_items):
                unclassified_spent += amt

        overall_exec_rate = round((total_spent / total_budget * 100), 1) if total_budget > 0 else 0.0
        overall_forecast_exec_rate = round((total_forecast_spent / total_budget * 100), 1) if total_budget > 0 else 0.0

        monthly_matrix = cls.calculate_monthly_matrix(budget_items, transactions)
        
        # 추경 변동 현황표 생성
        supp_matrix = {}
        if supplementary_items:
            supp_matrix = cls.calculate_supplementary_matrix(supplementary_items, spent_map)

        return {
            "remaining_months": remaining_months,
            "items": summary_items,
            "account_groups": list(account_groups.values()),
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
