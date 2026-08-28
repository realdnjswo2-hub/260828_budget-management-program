"""
e호조 23310 연동 세목별 예산관리대장 및 연말(12.31) 지출·잔액 예측 데스크톱 프로그램
- 당해년도 본예산 세출예산명세서 PDF 업로드 및 [단위-세부사업-편성목-통계목-세목] 5단계 계층 자동 인식/등록
- 1~4회 추경(추가경정예산) 세출사업명세서 PDF 업로드, 파싱 및 검토/승인 다이얼로그
- 1~4회 추경 변동 현황표 전용 탭 가시화
- 엑셀 업로드 시 지출 적요 패턴 기반 자동분류 규칙 자동 추출 및 학습
- 수동 규칙 추가/수정/삭제 팝업 모달 및 지출내역 더블클릭 연동 지원
- 규칙 내보내기/가져오기 백업 및 기본 규칙 초기화 기능
- 1월~12월 월별 지출 예산 통계 매트릭스 표 뷰 탑재
"""

import os
import sys
import json
import re
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from engine.parser import EHojoParser
from engine.pdf_parser import BaseBudgetPdfParser, SupplementaryPdfParser
from engine.classifier import KeywordClassifier
from engine.forecaster import Forecaster
from engine.excel_exporter import ExcelExporter


class BaseBudgetReviewDialog(tk.Toplevel):
    """본예산 세출예산명세서 PDF 파싱 결과 검토 및 승인 팝업 대화상자"""
    def __init__(self, parent, parsed_data):
        super().__init__(parent)
        self.parent = parent
        self.parsed_data = parsed_data
        self.result = None

        self.title(f"📘 {parsed_data.get('year', 2026)}년도 본예산 세출예산명세서 검토 및 등록")
        self.geometry("1060x620")
        self.minsize(880, 500)
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#F4F6F9")

        self._build_ui()
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def _build_ui(self):
        container = tk.Frame(self, bg="#FFFFFF", padx=16, pady=16, relief=tk.RAISED, bd=1)
        container.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        # 상단 타이틀 바
        top_bar = tk.Frame(container, bg="#FFFFFF")
        top_bar.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))

        title_text = f"📘 {self.parsed_data.get('title', '본예산 세출예산명세서')}"
        lbl_title = tk.Label(top_bar, text=title_text, font=("맑은 고딕", 12, "bold"), bg="#FFFFFF", fg="#1F497D")
        lbl_title.pack(side=tk.LEFT)

        total_amt = self.parsed_data.get("total_budget", 0)
        items_cnt = len(self.parsed_data.get("items", []))
        lbl_sum = tk.Label(top_bar, text=f"총 {items_cnt}개 세목  |  본예산 총액: {total_amt:,}원", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF", fg="#2E75B6")
        lbl_sum.pack(side=tk.RIGHT)

        # 안내문
        lbl_guide = tk.Label(
            container,
            text="💡 PDF 명세서에서 [단위사업 - 세부사업 - 편성목 - 통계목 - 세목(산출기초)] 및 예산액이 자동 추출되었습니다.\n확인 후 [당해년도 본예산 마스터로 확정 등록]을 누르면 예산대장 및 자동분류 규칙이 즉시 동기화됩니다.",
            font=("맑은 고딕", 9),
            bg="#EBF1F5",
            fg="#1E4E79",
            padx=10,
            pady=6,
            justify="left",
            relief=tk.GROOVE
        )
        lbl_guide.pack(fill=tk.X, side=tk.TOP, pady=(0, 8))

        # 테이블
        cols = ("unit_project", "detail_project", "category", "account", "sub_account", "budget", "calc_basis")
        self.tree = ttk.Treeview(container, columns=cols, show="headings", height=13)
        self.tree.heading("unit_project", text="단위사업")
        self.tree.heading("detail_project", text="세부사업")
        self.tree.heading("category", text="편성목")
        self.tree.heading("account", text="통계목")
        self.tree.heading("sub_account", text="세목 (산출기초)")
        self.tree.heading("budget", text="본예산액 (원)")
        self.tree.heading("calc_basis", text="산출기초 및 내역")

        self.tree.column("unit_project", width=140)
        self.tree.column("detail_project", width=140)
        self.tree.column("category", width=110)
        self.tree.column("account", width=130)
        self.tree.column("sub_account", width=150)
        self.tree.column("budget", width=110, anchor="e")
        self.tree.column("calc_basis", width=240)

        scroll_y = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # 데이터 추가
        for it in self.parsed_data.get("items", []):
            amt = int(it.get("budget", 0))
            self.tree.insert("", tk.END, values=(
                it.get("unit_project", ""),
                it.get("detail_project", ""),
                it.get("category", ""),
                it.get("account", ""),
                it.get("sub_account", ""),
                f"{amt:,}",
                it.get("calculation_basis", "")
            ))

        # 하단 버튼
        btn_bar = tk.Frame(container, bg="#FFFFFF")
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(12, 0))

        btn_apply = ttk.Button(btn_bar, text="💾 당해년도 본예산 마스터로 확정 등록", command=self._on_apply, style="Primary.TButton")
        btn_apply.pack(side=tk.LEFT, padx=6)

        btn_cancel = ttk.Button(btn_bar, text="취소", command=self.destroy, style="Action.TButton")
        btn_cancel.pack(side=tk.LEFT, padx=6)

    def _on_apply(self):
        items = self.parsed_data.get("items", [])
        if not items:
            messagebox.showwarning("데이터 없음", "등록할 본예산 항목이 없습니다.", parent=self)
            return

        budget_master = []
        supp_items = []
        inferred_rules = []

        for it in items:
            unit_p = it.get("unit_project", "기본행정 지원")
            det_p = it.get("detail_project", "부서 기본운영경비")
            cat = it.get("category", "물건비")
            acc = it.get("account", "기타")
            sub_acc = it.get("sub_account", "기타")
            b_amt = int(it.get("budget", 0))
            calc_str = it.get("calculation_basis", "")

            # 1. Budget Master 엔트리
            budget_master.append({
                "policy_project": it.get("policy_project", "일반행정"),
                "unit_project": unit_p,
                "detail_project": det_p,
                "category": cat,
                "account": acc,
                "sub_account": sub_acc,
                "budget": b_amt,
                "note": calc_str
            })

            # 2. Supplementary Budget 엔트리
            supp_items.append({
                "unit_project": unit_p,
                "detail_project": det_p,
                "account": acc,
                "sub_account": sub_acc,
                "base_budget": b_amt,
                "supplements": {"1": 0, "2": 0, "3": 0, "4": 0},
                "reasons": {},
                "final_budget": b_amt
            })

            # 3. 세목명 기반 자동분류 규칙 도출
            clean_sub = re.sub(r'[\(\)\[\]\<\>\,\.\;\:\'\"]', ' ', sub_acc).split()
            valid_kws = [k for k in clean_sub if len(k) >= 2 and k not in ["구입", "지급", "납부", "지원", "관리", "운영"]]
            if valid_kws:
                cond_str = " OR ".join(valid_kws)
                inferred_rules.append({
                    "target_account": acc,
                    "target_sub_account": sub_acc,
                    "condition": cond_str,
                    "priority": 15,
                    "auto_generated": True
                })

        self.result = {
            "budget_master": budget_master,
            "supplementary_budgets": supp_items,
            "inferred_rules": inferred_rules
        }
        self.destroy()


class SupplementaryReviewDialog(tk.Toplevel):
    """추경 PDF 파싱 결과 검토 및 승인 팝업 대화상자"""
    def __init__(self, parent, parsed_data, current_budget_master, current_supp_items):
        super().__init__(parent)
        self.parent = parent
        self.parsed_data = parsed_data
        self.current_budget_master = current_budget_master
        self.current_supp_items = current_supp_items
        self.result = None

        self.title("📄 추경 세출사업명세서 파싱 결과 검토 및 예산 반영")
        self.geometry("960x600")
        self.minsize(800, 480)
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#F4F6F9")

        self._build_ui()
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def _build_ui(self):
        container = tk.Frame(self, bg="#FFFFFF", padx=16, pady=16, relief=tk.RAISED, bd=1)
        container.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        top_bar = tk.Frame(container, bg="#FFFFFF")
        top_bar.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))

        tk.Label(top_bar, text="추경 차수 선택:", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 6))
        self.combo_round = ttk.Combobox(top_bar, values=["제1회 추경", "제2회 추경", "제3회 추경", "제4회 추경"], font=("맑은 고딕", 9, "bold"), width=12)
        self.combo_round.pack(side=tk.LEFT, padx=(0, 12))
        
        detected_round = self.parsed_data.get("round", 1)
        self.combo_round.current(max(0, min(3, detected_round - 1)))

        lbl_doc_title = tk.Label(top_bar, text=self.parsed_data.get("title", "추경 세출사업명세서"), font=("맑은 고딕", 11, "bold"), bg="#FFFFFF", fg="#1F497D")
        lbl_doc_title.pack(side=tk.LEFT)

        lbl_guide = tk.Label(
            container,
            text="💡 PDF 문서에서 추출된 세목별 추경 증감 내역입니다. 확인 후 [예산 마스터에 최종 반영]을 누르면 예산액과 추경 이력이 갱신됩니다.",
            font=("맑은 고딕", 9),
            bg="#EBF1F5",
            fg="#1E4E79",
            padx=8,
            pady=4,
            relief=tk.GROOVE
        )
        lbl_guide.pack(fill=tk.X, side=tk.TOP, pady=(0, 8))

        cols = ("account", "sub_account", "change", "reason")
        self.tree = ttk.Treeview(container, columns=cols, show="headings", height=12)
        self.tree.heading("account", text="통계목")
        self.tree.heading("sub_account", text="세목 (산출기초)")
        self.tree.heading("change", text="추경 증감액 (±원)")
        self.tree.heading("reason", text="증감 사유 및 산출근거")

        self.tree.column("account", width=140)
        self.tree.column("sub_account", width=180)
        self.tree.column("change", width=130, anchor="e")
        self.tree.column("reason", width=380)

        scroll_y = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("plus_item", foreground="#0070C0", font=("맑은 고딕", 9, "bold"))
        self.tree.tag_configure("minus_item", foreground="#C00000", font=("맑은 고딕", 9, "bold"))

        for it in self.parsed_data.get("items", []):
            amt = int(it.get("change_amount", 0))
            amt_str = f"+{amt:,}" if amt > 0 else f"{amt:,}"
            tag = "plus_item" if amt > 0 else ("minus_item" if amt < 0 else "")
            self.tree.insert("", tk.END, values=(
                it.get("account"),
                it.get("sub_account"),
                amt_str,
                it.get("reason", "")
            ), tags=(tag,) if tag else ())

        btn_bar = tk.Frame(container, bg="#FFFFFF")
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(12, 0))

        btn_apply = ttk.Button(btn_bar, text="💾 예산 마스터 및 추경 이력에 최종 반영", command=self._on_apply, style="Primary.TButton")
        btn_apply.pack(side=tk.LEFT, padx=6)

        btn_cancel = ttk.Button(btn_bar, text="취소", command=self.destroy, style="Action.TButton")
        btn_cancel.pack(side=tk.LEFT, padx=6)

    def _on_apply(self):
        round_idx = self.combo_round.current() + 1
        round_str = str(round_idx)

        supp_map = {}
        for s in self.current_supp_items:
            key = (s.get("account", "").strip(), s.get("sub_account", "").strip())
            supp_map[key] = dict(s)

        items_to_apply = self.parsed_data.get("items", [])
        if not items_to_apply:
            messagebox.showwarning("항목 없음", "반영할 추경 내역이 없습니다.", parent=self)
            return

        for it in items_to_apply:
            acc = it.get("account", "").strip()
            sub_acc = it.get("sub_account", "").strip()
            change = int(it.get("change_amount", 0))
            reason = it.get("reason", "")

            matched_key = None
            for (k_acc, k_sub) in supp_map.keys():
                if k_acc.split()[0] == acc.split()[0] and (k_sub in sub_acc or sub_acc in k_sub):
                    matched_key = (k_acc, k_sub)
                    break

            if not matched_key:
                matched_key = (acc, sub_acc)
                if matched_key not in supp_map:
                    supp_map[matched_key] = {
                        "unit_project": "기본행정 지원",
                        "detail_project": "부서 기본운영경비",
                        "account": acc,
                        "sub_account": sub_acc,
                        "base_budget": 0,
                        "supplements": {"1": 0, "2": 0, "3": 0, "4": 0},
                        "reasons": {},
                        "final_budget": 0
                    }

            entry = supp_map[matched_key]
            if "supplements" not in entry:
                entry["supplements"] = {"1": 0, "2": 0, "3": 0, "4": 0}
            if "reasons" not in entry:
                entry["reasons"] = {}

            entry["supplements"][round_str] = change
            if reason:
                entry["reasons"][round_str] = reason

            base_b = int(entry.get("base_budget", 0))
            tot_supp = sum(int(v) for v in entry["supplements"].values())
            entry["final_budget"] = base_b + tot_supp

        updated_budget_master = []
        for (acc, sub_acc), s_data in supp_map.items():
            updated_budget_master.append({
                "unit_project": s_data.get("unit_project", "기본행정 지원"),
                "detail_project": s_data.get("detail_project", "부서 기본운영경비"),
                "account": acc,
                "sub_account": sub_acc,
                "budget": s_data["final_budget"],
                "note": f"제{round_idx}회 추경 반영 (최종: {s_data['final_budget']:,}원)"
            })

        self.result = {
            "round": round_idx,
            "updated_supp_items": list(supp_map.values()),
            "updated_budget_master": updated_budget_master
        }
        self.destroy()


class RuleEditDialog(tk.Toplevel):
    """규칙 추가 및 수정을 위한 팝업 대화상자"""
    def __init__(self, parent, rule_data=None, account_options=None, initial_summary=""):
        super().__init__(parent)
        self.parent = parent
        self.rule_data = rule_data
        self.result = None

        self.title("✏ 자동분류 키워드 규칙 설정" if rule_data else "➕ 새 자동분류 키워드 규칙 추가")
        self.geometry("640x440")
        self.minsize(580, 400)
        self.transient(parent)
        self.grab_set()

        self.configure(bg="#F4F6F9")

        self.account_options = account_options or [
            "201-01 사무관리비", "201-02 공공운영비", "201-03 행사운영비",
            "202-01 국내여비", "202-03 국외여비",
            "203-01 기관운영업무추진비", "203-02 시책추진업무추진비", "203-04 부서운영업무추진비",
            "206-01 재료비", "307-02 민간경상사업보조", "401-01 시설비", "405-01 자산취득비"
        ]

        self._build_ui(initial_summary)
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def _build_ui(self, initial_summary):
        container = tk.Frame(self, bg="#FFFFFF", padx=16, pady=16, relief=tk.RAISED, bd=1)
        container.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        tk.Label(container, text="1. 통계목 선택:", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.combo_account = ttk.Combobox(container, values=self.account_options, font=("맑은 고딕", 9), width=35)
        self.combo_account.grid(row=0, column=1, sticky="w", pady=(0, 4))
        if self.rule_data:
            self.combo_account.set(self.rule_data.get("target_account", ""))
        else:
            self.combo_account.current(0)

        tk.Label(container, text="2. 매핑할 세목명:", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF").grid(row=1, column=0, sticky="w", pady=8)
        self.entry_sub_acc = tk.Entry(container, font=("맑은 고딕", 10), width=36)
        self.entry_sub_acc.grid(row=1, column=1, sticky="w", pady=8)
        if self.rule_data:
            self.entry_sub_acc.insert(0, self.rule_data.get("target_sub_account", ""))

        tk.Label(container, text="3. 키워드 조건식:", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF").grid(row=2, column=0, sticky="nw", pady=8)
        cond_frame = tk.Frame(container, bg="#FFFFFF")
        cond_frame.grid(row=2, column=1, sticky="we", pady=8)

        self.entry_condition = tk.Entry(cond_frame, font=("Consolas", 10), width=45)
        self.entry_condition.pack(fill=tk.X, side=tk.TOP, ipady=3)

        if self.rule_data:
            self.entry_condition.insert(0, self.rule_data.get("condition", ""))
        elif initial_summary:
            clean_kw = re.sub(r'[\(\)\[\]\<\>\,\.\;\:\'\"]', ' ', initial_summary).split()
            valid_kw = [k for k in clean_kw if len(k) >= 2][:4]
            self.entry_condition.insert(0, " OR ".join(valid_kw) if valid_kw else initial_summary[:15])

        btn_helper_frame = tk.Frame(cond_frame, bg="#FFFFFF")
        btn_helper_frame.pack(fill=tk.X, side=tk.TOP, pady=(6, 0))

        tk.Label(btn_helper_frame, text="입력 보조:", font=("맑은 고딕", 8), bg="#FFFFFF", fg="#666666").pack(side=tk.LEFT, padx=(0, 4))
        for op in [" OR ", " AND ", " AND NOT ", "(", ")"]:
            btn_op = tk.Button(
                btn_helper_frame, text=op.strip() or op, font=("Consolas", 8, "bold"), bg="#EAEFF5", relief=tk.GROOVE,
                command=lambda text=op: self._insert_op(text)
            )
            btn_op.pack(side=tk.LEFT, padx=2)

        lbl_help = tk.Label(
            container,
            text="예시 문법:\n• 복사용지 OR 토너 OR 잉크\n• (출장 OR 여비) AND NOT (세종 OR 서울)\n• (다과 OR 음료) AND 간담회",
            font=("맑은 고딕", 9),
            bg="#F9FAFC",
            fg="#4E5969",
            justify="left",
            padx=8,
            pady=6,
            relief=tk.SUNKEN
        )
        lbl_help.grid(row=3, column=0, columnspan=2, sticky="we", pady=12)

        btn_box = tk.Frame(container, bg="#FFFFFF")
        btn_box.grid(row=4, column=0, columnspan=2, pady=(10, 0))

        btn_save = ttk.Button(btn_box, text="💾 규칙 저장 및 즉시 반영", command=self._on_save, style="Primary.TButton")
        btn_save.pack(side=tk.LEFT, padx=6)

        btn_cancel = ttk.Button(btn_box, text="취소", command=self.destroy, style="Action.TButton")
        btn_cancel.pack(side=tk.LEFT, padx=6)

    def _insert_op(self, op_text):
        self.entry_condition.insert(tk.INSERT, op_text)
        self.entry_condition.focus_set()

    def _on_save(self):
        acc = self.combo_account.get().strip()
        sub_acc = self.entry_sub_acc.get().strip()
        cond = self.entry_condition.get().strip()

        if not acc:
            messagebox.showwarning("입력 오류", "통계목을 선택하세요.", parent=self)
            return
        if not sub_acc:
            messagebox.showwarning("입력 오류", "매핑할 세목명을 입력하세요.", parent=self)
            return
        if not cond:
            messagebox.showwarning("입력 오류", "키워드 조건식을 입력하세요.", parent=self)
            return

        self.result = {
            "target_account": acc,
            "target_sub_account": sub_acc,
            "condition": cond,
            "priority": 20,
            "auto_generated": False
        }
        self.destroy()


class BudgetApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("e호조(23310) 세출예산 세목별 예산관리대장 및 연말 지출예측 시스템")
        self.geometry("1460x880")
        self.minsize(1160, 760)

        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_dir = os.path.join(self.base_dir, "config")
        self.output_dir = os.path.join(self.base_dir, "output")
        os.makedirs(self.config_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)

        self.last_output_file = os.path.join(self.output_dir, f"{datetime.now().year}년_세출예산_세목별_예산관리대장.xlsx")
        self.selected_file_path = tk.StringVar(value="")

        self.budget_master = []
        self.supplementary_budgets = []
        self.rules = []
        self.recurring_plans = []
        self.scheduled_plans = []
        self.transactions = []
        self.raw_transactions = []
        self.simulation_result = {}

        self.classifier = KeywordClassifier()

        self._setup_styles()
        self._load_configs()
        self._build_ui()

        self.log("시스템 준비 완료: 1단계 [📘 본예산 명세서(PDF) 등록] 또는 3단계 [▶ e호조 분석 시작]을 이용하세요.", "INFO")

    def _setup_styles(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.configure(bg="#F4F6F9")

        self.style.configure("Primary.TButton", font=("맑은 고딕", 10, "bold"), background="#1E4E79", foreground="#FFFFFF")
        self.style.map("Primary.TButton", background=[("active", "#15375B"), ("pressed", "#0F263E")])

        self.style.configure("BaseBudget.TButton", font=("맑은 고딕", 10, "bold"), background="#0E6251", foreground="#FFFFFF")
        self.style.map("BaseBudget.TButton", background=[("active", "#0B4C3E")])

        self.style.configure("Purple.TButton", font=("맑은 고딕", 10, "bold"), background="#4B2C82", foreground="#FFFFFF")
        self.style.map("Purple.TButton", background=[("active", "#361F5E")])

        self.style.configure("Success.TButton", font=("맑은 고딕", 10, "bold"), background="#2E75B6", foreground="#FFFFFF")
        self.style.map("Success.TButton", background=[("active", "#1F4E79")])

        self.style.configure("Action.TButton", font=("맑은 고딕", 9), background="#D9E1F2", foreground="#1F497D")
        self.style.map("Action.TButton", background=[("active", "#B4C6E7")])

        self.style.configure("AI.TButton", font=("맑은 고딕", 9, "bold"), background="#4B2C82", foreground="#FFFFFF")
        self.style.map("AI.TButton", background=[("active", "#361F5E")])

        self.style.configure(
            "Budget.Treeview",
            font=("맑은 고딕", 9),
            rowheight=25,
            background="#FFFFFF",
            fieldbackground="#FFFFFF"
        )
        self.style.configure(
            "Budget.Treeview.Heading",
            font=("맑은 고딕", 9, "bold"),
            background="#2F4F4F",
            foreground="#FFFFFF",
            padding=4
        )
        self.style.map("Budget.Treeview.Heading", background=[("active", "#1D3232")])

    def _load_configs(self):
        bm_path = os.path.join(self.config_dir, "budget_master.json")
        if os.path.exists(bm_path):
            with open(bm_path, "r", encoding="utf-8") as f:
                self.budget_master = json.load(f)
        else:
            self.budget_master = []

        supp_path = os.path.join(self.config_dir, "supplementary_budgets.json")
        if os.path.exists(supp_path):
            with open(supp_path, "r", encoding="utf-8") as f:
                self.supplementary_budgets = json.load(f)
        else:
            self.supplementary_budgets = []

        rules_path = os.path.join(self.config_dir, "rules.json")
        if os.path.exists(rules_path):
            with open(rules_path, "r", encoding="utf-8") as f:
                self.rules = json.load(f)
        else:
            self.rules = KeywordClassifier.get_default_rules()
            self._save_configs()
            
        self.classifier.set_rules(self.rules)

        rec_path = os.path.join(self.config_dir, "recurring_plans.json")
        if os.path.exists(rec_path):
            with open(rec_path, "r", encoding="utf-8") as f:
                self.recurring_plans = json.load(f)
        else:
            self.recurring_plans = []

        sched_path = os.path.join(self.config_dir, "scheduled_plans.json")
        if os.path.exists(sched_path):
            with open(sched_path, "r", encoding="utf-8") as f:
                self.scheduled_plans = json.load(f)
        else:
            self.scheduled_plans = []

    def _save_configs(self):
        with open(os.path.join(self.config_dir, "budget_master.json"), "w", encoding="utf-8") as f:
            json.dump(self.budget_master, f, ensure_ascii=False, indent=2)
        with open(os.path.join(self.config_dir, "supplementary_budgets.json"), "w", encoding="utf-8") as f:
            json.dump(self.supplementary_budgets, f, ensure_ascii=False, indent=2)
        with open(os.path.join(self.config_dir, "rules.json"), "w", encoding="utf-8") as f:
            json.dump(self.rules, f, ensure_ascii=False, indent=2)
        with open(os.path.join(self.config_dir, "recurring_plans.json"), "w", encoding="utf-8") as f:
            json.dump(self.recurring_plans, f, ensure_ascii=False, indent=2)
        with open(os.path.join(self.config_dir, "scheduled_plans.json"), "w", encoding="utf-8") as f:
            json.dump(self.scheduled_plans, f, ensure_ascii=False, indent=2)

    def _build_ui(self):
        top_frame = tk.Frame(self, bg="#FFFFFF", padx=14, pady=10, relief=tk.RAISED, bd=1)
        top_frame.pack(fill=tk.X, side=tk.TOP)

        title_lbl = tk.Label(
            top_frame,
            text="📊 e호조(23310) 세출예산 세목별 예산관리대장 및 연말 지출예측 시스템",
            font=("맑은 고딕", 13, "bold"),
            bg="#FFFFFF",
            fg="#1F497D"
        )
        title_lbl.pack(side=tk.TOP, anchor="w", pady=(0, 6))

        btn_bar = tk.Frame(top_frame, bg="#FFFFFF")
        btn_bar.pack(fill=tk.X, side=tk.TOP)

        # 1단계: 📘 본예산 명세서(PDF) 등록
        btn_base_pdf = ttk.Button(btn_bar, text="📘 1단계: 본예산 명세서(PDF) 등록", command=self._on_upload_base_budget_pdf, style="BaseBudget.TButton")
        btn_base_pdf.pack(side=tk.LEFT, padx=(0, 6), ipady=1)

        # 2단계: 📄 추경 사업명세서(PDF) 업로드
        btn_upload_pdf = ttk.Button(btn_bar, text="📄 2단계: 추경 명세서(PDF) 업로드", command=self._on_upload_supplementary_pdf, style="Purple.TButton")
        btn_upload_pdf.pack(side=tk.LEFT, padx=(0, 6), ipady=1)

        lbl_file = tk.Label(btn_bar, text="e호조 파일:", font=("맑은 고딕", 9, "bold"), bg="#FFFFFF")
        lbl_file.pack(side=tk.LEFT, padx=(6, 4))

        self.entry_file = tk.Entry(btn_bar, textvariable=self.selected_file_path, width=28, font=("맑은 고딕", 9))
        self.entry_file.pack(side=tk.LEFT, padx=(0, 4), ipady=2)

        btn_browse = ttk.Button(btn_bar, text="📁 찾기", command=self._on_browse_file, style="Action.TButton")
        btn_browse.pack(side=tk.LEFT, padx=(0, 6))

        # 3단계: ▶ e호조 지출 분석 시작
        btn_analyze = ttk.Button(btn_bar, text="▶ 3단계: e호조 지출 분석 시작", command=self._on_start_analysis, style="Primary.TButton")
        btn_analyze.pack(side=tk.LEFT, padx=(0, 6), ipady=1)

        btn_open_excel = ttk.Button(btn_bar, text="📊 엑셀 열기", command=self._on_open_excel, style="Success.TButton")
        btn_open_excel.pack(side=tk.LEFT, padx=(0, 6), ipady=1)

        btn_open_folder = ttk.Button(btn_bar, text="📂 저장 폴더", command=self._on_open_folder, style="Action.TButton")
        btn_open_folder.pack(side=tk.LEFT, padx=(0, 6), ipady=1)

        btn_sample = ttk.Button(btn_bar, text="💡 샘플 파일 생성", command=self._on_generate_sample, style="Action.TButton")
        btn_sample.pack(side=tk.RIGHT, padx=2)

        # =========================================================================
        # 중앙 탭 노트북 (Tab View)
        # =========================================================================
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=(8, 4))

        # 탭 1: 세목별 예산관리대장
        self.tab_ledger = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab_ledger, text="  📋 세목별 예산관리대장 (연말예측)  ")
        self._build_ledger_tab()

        # 탭 2: 📊 추경 예산 변동 현황표 (1~4회)
        self.tab_supp = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab_supp, text="  📊 추경 예산 변동 현황표 (1~4회)  ")
        self._build_supp_tab()

        # 탭 3: 📅 월별 지출 예산 통계표 (1~12월)
        self.tab_monthly = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab_monthly, text="  📅 월별 지출 예산 통계표 (1~12월)  ")
        self._build_monthly_tab()

        # 탭 4: e호조 수집 지출내역
        self.tab_transactions = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab_transactions, text="  📝 e호조 지출 수집 내역  ")
        self._build_transactions_tab()

        # 탭 5: 12.31 연말 계획 설정
        self.tab_forecast = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab_forecast, text="  🔮 연말 계획 설정  ")
        self._build_forecast_tab()

        # 탭 6: 규칙 및 세목 마스터 설정
        self.tab_rules = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(self.tab_rules, text="  ⚙ 세목 & 키워드 자동분류 규칙  ")
        self._build_rules_tab()

        # =========================================================================
        # 하단 진행 로그 창
        # =========================================================================
        log_frame = tk.LabelFrame(
            self,
            text=" 📝 진행 로그 (실시간 상태) ",
            font=("맑은 고딕", 9, "bold"),
            bg="#F4F6F9",
            fg="#333333",
            padx=8,
            pady=4
        )
        log_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=12, pady=(0, 8))

        self.log_text = tk.Text(log_frame, height=5, font=("Consolas", 9), bg="#1E1E1E", fg="#D4D4D4", relief=tk.FLAT)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text.tag_config("INFO", foreground="#4EC9B0")
        self.log_text.tag_config("SUCCESS", foreground="#B5CEA8")
        self.log_text.tag_config("WARN", foreground="#CE9178")
        self.log_text.tag_config("ERROR", foreground="#F44747")

    # -------------------------------------------------------------------------
    # 탭 1: 세목별 예산관리대장
    # -------------------------------------------------------------------------
    def _build_ledger_tab(self):
        info_frame = tk.Frame(self.tab_ledger, bg="#EEF2F7", padx=10, pady=6)
        info_frame.pack(fill=tk.X, side=tk.TOP)

        self.lbl_ledger_summary = tk.Label(
            info_frame,
            text="총 배정예산: 0원  |  총 기집행액: 0원 (집행률: 0.0%)  |  현재잔액: 0원  |  12.31 예상잔액: 0원",
            font=("맑은 고딕", 10, "bold"),
            bg="#EEF2F7",
            fg="#1F497D"
        )
        self.lbl_ledger_summary.pack(side=tk.LEFT)

        cols = (
            "detail_project", "category", "account", "sub_account", "budget", "actual_spent", "current_balance",
            "exec_rate", "remaining_recurring", "scheduled_spent", "forecast_total_spent",
            "forecast_balance", "status"
        )
        self.tree_ledger = ttk.Treeview(self.tab_ledger, columns=cols, show="headings", style="Budget.Treeview")

        headings = [
            ("detail_project", "세부사업", 130, "w"),
            ("category", "편성목", 100, "w"),
            ("account", "통계목", 125, "w"),
            ("sub_account", "세목 (산출기초)", 150, "w"),
            ("budget", "배정예산액 (A)", 105, "e"),
            ("actual_spent", "기집행액 (B)", 100, "e"),
            ("current_balance", "현재잔액 (A-B)", 100, "e"),
            ("exec_rate", "집행률", 70, "center"),
            ("remaining_recurring", "잔여정기지출", 90, "e"),
            ("scheduled_spent", "하반기예정액", 90, "e"),
            ("forecast_total_spent", "12.31 예상지출(C)", 110, "e"),
            ("forecast_balance", "12.31 예상잔액(A-C)", 110, "e"),
            ("status", "상태판정", 95, "center")
        ]

        for col_id, col_text, col_w, col_align in headings:
            self.tree_ledger.heading(col_id, text=col_text)
            self.tree_ledger.column(col_id, width=col_w, anchor=col_align)

        tree_scroll_y = ttk.Scrollbar(self.tab_ledger, orient=tk.VERTICAL, command=self.tree_ledger.yview)
        tree_scroll_x = ttk.Scrollbar(self.tab_ledger, orient=tk.HORIZONTAL, command=self.tree_ledger.xview)
        self.tree_ledger.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        self.tree_ledger.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=6)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y, pady=6, padx=(0, 8))
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X, padx=8)

        self.tree_ledger.tag_configure("header_row", background="#E8EEF5", font=("맑은 고딕", 9, "bold"))
        self.tree_ledger.tag_configure("total_row", background="#D9E1F2", font=("맑은 고딕", 9, "bold"))
        self.tree_ledger.tag_configure("danger_row", foreground="#C00000")
        self.tree_ledger.tag_configure("warning_row", foreground="#ED7D31")

    # -------------------------------------------------------------------------
    # 탭 2: 📊 추경 예산 변동 현황표 (1~4회)
    # -------------------------------------------------------------------------
    def _build_supp_tab(self):
        info_frame = tk.Frame(self.tab_supp, bg="#F3EEF9", padx=10, pady=6)
        info_frame.pack(fill=tk.X, side=tk.TOP)

        self.lbl_supp_summary = tk.Label(
            info_frame,
            text="📊 당초 본예산 대비 1~4회 추가경정예산 증감 내역 및 최종 확정예산 대비 집행 현황",
            font=("맑은 고딕", 10, "bold"),
            bg="#F3EEF9",
            fg="#4B2C82"
        )
        self.lbl_supp_summary.pack(side=tk.LEFT)

        cols = (
            "detail_project", "account", "sub_account", "base_budget",
            "r1", "r2", "r3", "r4",
            "final_budget", "spent", "balance", "exec_rate", "reason"
        )
        self.tree_supp = ttk.Treeview(self.tab_supp, columns=cols, show="headings", style="Budget.Treeview")

        s_headings = [
            ("detail_project", "세부사업", 120, "w"),
            ("account", "통계목", 120, "w"),
            ("sub_account", "세목 (산출기초)", 140, "w"),
            ("base_budget", "당초 본예산", 95, "e"),
            ("r1", "1회 추경(±)", 85, "e"),
            ("r2", "2회 추경(±)", 85, "e"),
            ("r3", "3회 추경(±)", 85, "e"),
            ("r4", "4회 추경(±)", 85, "e"),
            ("final_budget", "최종 확정예산", 100, "e"),
            ("spent", "기집행액", 90, "e"),
            ("balance", "현재잔액", 90, "e"),
            ("exec_rate", "집행률", 60, "center"),
            ("reason", "증감 사유 및 내역", 230, "w")
        ]

        for col_id, col_text, col_w, col_align in s_headings:
            self.tree_supp.heading(col_id, text=col_text)
            self.tree_supp.column(col_id, width=col_w, anchor=col_align)

        s_scroll_y = ttk.Scrollbar(self.tab_supp, orient=tk.VERTICAL, command=self.tree_supp.yview)
        s_scroll_x = ttk.Scrollbar(self.tab_supp, orient=tk.HORIZONTAL, command=self.tree_supp.xview)
        self.tree_supp.configure(yscrollcommand=s_scroll_y.set, xscrollcommand=s_scroll_x.set)

        self.tree_supp.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=6)
        s_scroll_y.pack(side=tk.RIGHT, fill=tk.Y, pady=6, padx=(0, 8))
        s_scroll_x.pack(side=tk.BOTTOM, fill=tk.X, padx=8)

        self.tree_supp.tag_configure("header_row", background="#E8EEF5", font=("맑은 고딕", 9, "bold"))
        self.tree_supp.tag_configure("total_row", background="#D9E1F2", font=("맑은 고딕", 9, "bold"))

    # -------------------------------------------------------------------------
    # 탭 3: 📅 월별 지출 예산 통계표
    # -------------------------------------------------------------------------
    def _build_monthly_tab(self):
        info_frame = tk.Frame(self.tab_monthly, bg="#F2F4F7", padx=10, pady=6)
        info_frame.pack(fill=tk.X, side=tk.TOP)

        self.lbl_monthly_summary = tk.Label(
            info_frame,
            text="📅 세부사업 및 통계목/세목별 월별(1월~12월) 지출 추이 및 누적 현황표",
            font=("맑은 고딕", 10, "bold"),
            bg="#F2F4F7",
            fg="#1E4E79"
        )
        self.lbl_monthly_summary.pack(side=tk.LEFT)

        cols = (
            "detail_project", "account", "sub_account", "budget",
            "m1", "m2", "m3", "m4", "m5", "m6", "m7", "m8", "m9", "m10", "m11", "m12",
            "total_spent", "balance", "exec_rate"
        )
        self.tree_monthly = ttk.Treeview(self.tab_monthly, columns=cols, show="headings", style="Budget.Treeview")

        m_headings = [
            ("detail_project", "세부사업", 120, "w"),
            ("account", "통계목", 120, "w"),
            ("sub_account", "세목 (산출기초)", 140, "w"),
            ("budget", "배정예산액", 90, "e"),
            ("m1", "1월", 65, "e"),
            ("m2", "2월", 65, "e"),
            ("m3", "3월", 65, "e"),
            ("m4", "4월", 65, "e"),
            ("m5", "5월", 65, "e"),
            ("m6", "6월", 65, "e"),
            ("m7", "7월", 65, "e"),
            ("m8", "8월", 65, "e"),
            ("m9", "9월", 65, "e"),
            ("m10", "10월", 65, "e"),
            ("m11", "11월", 65, "e"),
            ("m12", "12월", 65, "e"),
            ("total_spent", "누적집행액", 90, "e"),
            ("balance", "현재잔액", 90, "e"),
            ("exec_rate", "집행률", 60, "center")
        ]

        for col_id, col_text, col_w, col_align in m_headings:
            self.tree_monthly.heading(col_id, text=col_text)
            self.tree_monthly.column(col_id, width=col_w, anchor=col_align)

        m_scroll_y = ttk.Scrollbar(self.tab_monthly, orient=tk.VERTICAL, command=self.tree_monthly.yview)
        m_scroll_x = ttk.Scrollbar(self.tab_monthly, orient=tk.HORIZONTAL, command=self.tree_monthly.xview)
        self.tree_monthly.configure(yscrollcommand=m_scroll_y.set, xscrollcommand=m_scroll_x.set)

        self.tree_monthly.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=6)
        m_scroll_y.pack(side=tk.RIGHT, fill=tk.Y, pady=6, padx=(0, 8))
        m_scroll_x.pack(side=tk.BOTTOM, fill=tk.X, padx=8)

        self.tree_monthly.tag_configure("header_row", background="#E8EEF5", font=("맑은 고딕", 9, "bold"))
        self.tree_monthly.tag_configure("total_row", background="#D9E1F2", font=("맑은 고딕", 9, "bold"))

    # -------------------------------------------------------------------------
    # 탭 4: e호조 수집 지출내역
    # -------------------------------------------------------------------------
    def _build_transactions_tab(self):
        top_filter = tk.Frame(self.tab_transactions, bg="#FFFFFF", padx=10, pady=6)
        top_filter.pack(fill=tk.X, side=tk.TOP)

        tk.Label(top_filter, text="검색 (적요/채권자/세목):", font=("맑은 고딕", 9), bg="#FFFFFF").pack(side=tk.LEFT, padx=(0, 4))
        self.entry_search_tx = tk.Entry(top_filter, width=28, font=("맑은 고딕", 9))
        self.entry_search_tx.pack(side=tk.LEFT, padx=(0, 6))
        self.entry_search_tx.bind("<KeyRelease>", self._filter_transactions)

        btn_rule_from_sel = ttk.Button(top_filter, text="➕ 선택 내역으로 규칙 만들기", command=self._on_create_rule_from_selected_tx, style="Action.TButton")
        btn_rule_from_sel.pack(side=tk.LEFT, padx=6)

        tk.Label(top_filter, text="💡 행을 더블클릭하면 즉시 해당 적요로 규칙 생성 창이 열립니다.", font=("맑은 고딕", 8), bg="#FFFFFF", fg="#666666").pack(side=tk.LEFT, padx=6)

        self.lbl_tx_count = tk.Label(top_filter, text="총 0건의 지출 내역", font=("맑은 고딕", 9, "bold"), bg="#FFFFFF", fg="#1F497D")
        self.lbl_tx_count.pack(side=tk.RIGHT)

        cols = ("idx", "date", "account", "sub_account", "summary", "vendor", "amount", "rule")
        self.tree_tx = ttk.Treeview(self.tab_transactions, columns=cols, show="headings", style="Budget.Treeview")

        headings = [
            ("idx", "연번", 50, "center"),
            ("date", "결의일자", 95, "center"),
            ("account", "통계목", 130, "w"),
            ("sub_account", "자동분류 세목", 150, "w"),
            ("summary", "지출 적요 (온나라 기안명)", 360, "w"),
            ("vendor", "채권자 (지급처)", 130, "w"),
            ("amount", "지출액 (원)", 110, "e"),
            ("rule", "적용 키워드 규칙", 180, "w")
        ]

        for col_id, col_text, col_w, col_align in headings:
            self.tree_tx.heading(col_id, text=col_text)
            self.tree_tx.column(col_id, width=col_w, anchor=col_align)

        self.tree_tx.bind("<Double-1>", self._on_double_click_tx)

        scroll_y = ttk.Scrollbar(self.tab_transactions, orient=tk.VERTICAL, command=self.tree_tx.yview)
        scroll_x = ttk.Scrollbar(self.tab_transactions, orient=tk.HORIZONTAL, command=self.tree_tx.xview)
        self.tree_tx.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree_tx.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=6)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y, pady=6, padx=(0, 8))
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X, padx=8)

        self.tree_tx.tag_configure("unclassified", foreground="#ED7D31", font=("맑은 고딕", 9, "bold"))

    # -------------------------------------------------------------------------
    # 탭 5: 연말 계획 설정
    # -------------------------------------------------------------------------
    def _build_forecast_tab(self):
        container = tk.Frame(self.tab_forecast, bg="#FFFFFF", padx=10, pady=10)
        container.pack(fill=tk.BOTH, expand=True)

        left_frame = tk.LabelFrame(container, text=" 🔄 매월 반복 고정지출 (월정액 × 12월까지 잔여월수) ", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF", padx=8, pady=8)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        self.tree_rec = ttk.Treeview(left_frame, columns=("account", "sub_account", "monthly", "desc"), show="headings", height=10)
        self.tree_rec.heading("account", text="통계목")
        self.tree_rec.heading("sub_account", text="세목")
        self.tree_rec.heading("monthly", text="월정액 (원)")
        self.tree_rec.heading("desc", text="지출 성격")
        self.tree_rec.column("account", width=120)
        self.tree_rec.column("sub_account", width=130)
        self.tree_rec.column("monthly", width=95, anchor="e")
        self.tree_rec.column("desc", width=160)
        self.tree_rec.pack(fill=tk.BOTH, expand=True)

        right_frame = tk.LabelFrame(container, text=" 📌 하반기(9~12월) 집행예정 사업비 계획 ", font=("맑은 고딕", 10, "bold"), bg="#FFFFFF", padx=8, pady=8)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))

        self.tree_sched = ttk.Treeview(right_frame, columns=("account", "sub_account", "month", "amount", "desc"), show="headings", height=10)
        self.tree_sched.heading("account", text="통계목")
        self.tree_sched.heading("sub_account", text="세목")
        self.tree_sched.heading("month", text="예정월")
        self.tree_sched.heading("amount", text="예정금액 (원)")
        self.tree_sched.heading("desc", text="사업 내용")
        self.tree_sched.column("account", width=110)
        self.tree_sched.column("sub_account", width=120)
        self.tree_sched.column("month", width=60, anchor="center")
        self.tree_sched.column("amount", width=95, anchor="e")
        self.tree_sched.column("desc", width=160)
        self.tree_sched.pack(fill=tk.BOTH, expand=True)

        self._render_forecast_tab_data()

    def _render_forecast_tab_data(self):
        for item in self.tree_rec.get_children():
            self.tree_rec.delete(item)
        for r in self.recurring_plans:
            self.tree_rec.insert("", tk.END, values=(
                r.get("account"), r.get("sub_account"), f"{r.get('monthly_amount', 0):,}", r.get("description")
            ))

        for item in self.tree_sched.get_children():
            self.tree_sched.delete(item)
        for s in self.scheduled_plans:
            self.tree_sched.insert("", tk.END, values=(
                s.get("account"), s.get("sub_account"), f"{s.get('month', 12)}월", f"{s.get('amount', 0):,}", s.get("description")
            ))

    # -------------------------------------------------------------------------
    # 탭 6: 규칙 및 세목 설정 빌드
    # -------------------------------------------------------------------------
    def _build_rules_tab(self):
        container = tk.Frame(self.tab_rules, bg="#FFFFFF", padx=10, pady=10)
        container.pack(fill=tk.BOTH, expand=True)

        top_rule_bar = tk.Frame(container, bg="#FFFFFF")
        top_rule_bar.pack(fill=tk.X, side=tk.TOP, pady=(0, 6))

        btn_add = ttk.Button(top_rule_bar, text="➕ 새 규칙 추가", command=self._on_add_rule_dialog, style="Primary.TButton")
        btn_add.pack(side=tk.LEFT, padx=(0, 4))

        btn_edit = ttk.Button(top_rule_bar, text="✏ 선택 규칙 수정", command=self._on_edit_rule_dialog, style="Action.TButton")
        btn_edit.pack(side=tk.LEFT, padx=4)

        btn_del = ttk.Button(top_rule_bar, text="🗑 선택 규칙 삭제", command=self._on_delete_rule, style="Action.TButton")
        btn_del.pack(side=tk.LEFT, padx=4)

        btn_export = ttk.Button(top_rule_bar, text="💾 규칙 내보내기", command=self._on_export_rules, style="Action.TButton")
        btn_export.pack(side=tk.LEFT, padx=4)

        btn_import = ttk.Button(top_rule_bar, text="📂 규칙 가져오기", command=self._on_import_rules, style="Action.TButton")
        btn_import.pack(side=tk.LEFT, padx=4)

        btn_reset = ttk.Button(top_rule_bar, text="🔄 기본 규칙 초기화", command=self._on_reset_rules, style="Action.TButton")
        btn_reset.pack(side=tk.LEFT, padx=4)

        btn_auto_infer = ttk.Button(
            top_rule_bar,
            text="🧠 지출내역 기반 규칙 자동 추출/학습",
            command=self._on_auto_infer_rules_manual,
            style="AI.TButton"
        )
        btn_auto_infer.pack(side=tk.RIGHT)

        cols = ("account", "sub_account", "condition", "priority", "auto")
        self.tree_rules_view = ttk.Treeview(container, columns=cols, show="headings", style="Budget.Treeview")
        self.tree_rules_view.heading("account", text="통계목")
        self.tree_rules_view.heading("sub_account", text="매핑 세목")
        self.tree_rules_view.heading("condition", text="키워드 매칭 조건식 (AND / OR / NOT)")
        self.tree_rules_view.heading("priority", text="우선순위")
        self.tree_rules_view.heading("auto", text="구분")

        self.tree_rules_view.column("account", width=140)
        self.tree_rules_view.column("sub_account", width=160)
        self.tree_rules_view.column("condition", width=500)
        self.tree_rules_view.column("priority", width=70, anchor="center")
        self.tree_rules_view.column("auto", width=90, anchor="center")

        self.tree_rules_view.bind("<Double-1>", lambda e: self._on_edit_rule_dialog())

        scroll_y = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.tree_rules_view.yview)
        self.tree_rules_view.configure(yscrollcommand=scroll_y.set)

        self.tree_rules_view.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self._render_rules_tab_data()

    def _render_rules_tab_data(self):
        for item in self.tree_rules_view.get_children():
            self.tree_rules_view.delete(item)
        for r in self.rules:
            is_auto = "✨ 자동생성" if r.get("auto_generated") else "수동작성"
            self.tree_rules_view.insert("", tk.END, values=(
                r.get("target_account"), r.get("target_sub_account"), r.get("condition"), r.get("priority", 10), is_auto
            ))

    # =========================================================================
    # 이벤트 핸들러 및 비즈니스 로직
    # =========================================================================
    def log(self, message: str, level="INFO"):
        now_str = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{now_str}] [{level}] {message}\n"
        self.log_text.insert(tk.END, log_entry, level)
        self.log_text.see(tk.END)

    def _on_browse_file(self):
        file_path = filedialog.askopenfilename(
            title="e호조 23310 엑셀/CSV 파일 선택",
            filetypes=[("엑셀 파일 (*.xlsx)", "*.xlsx"), ("CSV 파일 (*.csv)", "*.csv"), ("모든 파일", "*.*")]
        )
        if file_path:
            self.selected_file_path.set(file_path)
            self.log(f"파일 선택 완료: {os.path.basename(file_path)}", "INFO")

    def _on_generate_sample(self):
        sample_path = os.path.join(self.base_dir, "sample", "e호조_23310_샘플.xlsx")
        from sample_generator import generate_sample_file, generate_sample_base_budget_txt, generate_sample_supplementary_txt
        generate_sample_file(sample_path)
        generate_sample_base_budget_txt(os.path.join(self.base_dir, "sample", "2026년도_본예산_세출예산명세서_샘플.txt"))
        generate_sample_supplementary_txt(os.path.join(self.base_dir, "sample", "제1회_추경_세출사업명세서_샘플.txt"))
        self.selected_file_path.set(sample_path)
        self.log("테스트용 본예산 명세서, 추경 명세서 및 e호조 엑셀 샘플 파일이 생성되었습니다.", "SUCCESS")
        messagebox.showinfo("샘플 생성 완료", "테스트용 [본예산 명세서], [추경 명세서], [e호조 엑셀] 샘플 파일이\n'sample' 폴더에 생성되었습니다.")

    def _on_upload_base_budget_pdf(self):
        """1단계: 당해년도 본예산 세출예산명세서 PDF/TXT 업로드 및 파싱"""
        file_path = filedialog.askopenfilename(
            title="당해년도 본예산 세출예산명세서 PDF/TXT 선택",
            filetypes=[("PDF/TXT 파일 (*.pdf;*.txt)", "*.pdf;*.txt"), ("PDF 파일 (*.pdf)", "*.pdf"), ("텍스트 파일 (*.txt)", "*.txt"), ("모든 파일", "*.*")]
        )
        if not file_path:
            sample_txt = os.path.join(self.base_dir, "sample", "2026년도_본예산_세출예산명세서_샘플.txt")
            if os.path.exists(sample_txt):
                file_path = sample_txt
                self.log("선택된 파일이 없어 테스트용 본예산 샘플 명세서를 불러옵니다.", "WARN")
            else:
                return

        self.log(f"📘 본예산 세출예산명세서 분석 시작: {os.path.basename(file_path)}", "INFO")
        try:
            parsed = BaseBudgetPdfParser.parse_base_budget_pdf(file_path)
            self.log(f"본예산 파싱 완료: {parsed.get('title')} (추출된 세목: {len(parsed.get('items', []))}건, 총액: {parsed.get('total_budget', 0):,}원)", "SUCCESS")

            dialog = BaseBudgetReviewDialog(self, parsed_data=parsed)
            self.wait_window(dialog)

            if dialog.result:
                self.budget_master = dialog.result["budget_master"]
                self.supplementary_budgets = dialog.result["supplementary_budgets"]
                
                # 신규 규칙 자동 추가
                new_rules = dialog.result.get("inferred_rules", [])
                existing_conds = {r.get("condition", "") for r in self.rules}
                for nr in new_rules:
                    if nr.get("condition") not in existing_conds:
                        self.rules.append(nr)

                self._save_configs()
                self.classifier.set_rules(self.rules)
                self._render_rules_tab_data()

                self.log(f"🎉 {parsed.get('year', 2026)}년도 본예산 마스터 ({len(self.budget_master)}개 세목)가 성공적으로 확정 등록되었습니다!", "SUCCESS")
                self._reclassify_and_refresh()
                messagebox.showinfo(
                    "본예산 등록 완료",
                    f"{parsed.get('year', 2026)}년도 본예산 마스터가 성공적으로 등록되었습니다!\n\n"
                    f"• 등록 세목: {len(self.budget_master)}개\n"
                    f"• 본예산 총액: {parsed.get('total_budget', 0):,}원\n"
                    f"• 세목 기반 자동분류 규칙이 동기화되었습니다."
                )

        except Exception as e:
            self.log(f"본예산 파일 파싱 오류: {str(e)}", "ERROR")
            messagebox.showerror("파싱 오류", f"본예산 명세서 분석 중 오류가 발생했습니다:\n{str(e)}")

    def _on_upload_supplementary_pdf(self):
        """2단계: 추경 세출사업명세서 PDF/TXT 업로드 및 파싱"""
        file_path = filedialog.askopenfilename(
            title="추경 세출사업명세서 PDF/TXT 파일 선택",
            filetypes=[("PDF/TXT 파일 (*.pdf;*.txt)", "*.pdf;*.txt"), ("PDF 파일 (*.pdf)", "*.pdf"), ("텍스트 파일 (*.txt)", "*.txt"), ("모든 파일", "*.*")]
        )
        if not file_path:
            sample_txt = os.path.join(self.base_dir, "sample", "제1회_추경_세출사업명세서_샘플.txt")
            if os.path.exists(sample_txt):
                file_path = sample_txt
                self.log("선택된 파일이 없어 테스트용 제1회 추경 샘플 명세서를 불러옵니다.", "WARN")
            else:
                return

        self.log(f"📄 추경 사업명세서 분석 시작: {os.path.basename(file_path)}", "INFO")
        try:
            parsed = SupplementaryPdfParser.parse_supplementary_pdf(file_path)
            self.log(f"추경 명세서 파싱 완료: {parsed.get('title')} (추출된 항목: {len(parsed.get('items', []))}건)", "SUCCESS")
            
            dialog = SupplementaryReviewDialog(
                self,
                parsed_data=parsed,
                current_budget_master=self.budget_master,
                current_supp_items=self.supplementary_budgets
            )
            self.wait_window(dialog)

            if dialog.result:
                self.supplementary_budgets = dialog.result["updated_supp_items"]
                self.budget_master = dialog.result["updated_budget_master"]
                self._save_configs()
                
                self.log(f"🎉 제{dialog.result['round']}회 추경 예산이 성공적으로 마스터에 반영되었습니다!", "SUCCESS")
                self._reclassify_and_refresh()
                messagebox.showinfo("추경 예산 반영 완료", f"제{dialog.result['round']}회 추경 예산이 성공적으로 반영되었습니다!\n예산관리대장과 1~4회 추경 변동 현황표가 갱신되었습니다.")

        except Exception as e:
            self.log(f"추경 파일 파싱 오류: {str(e)}", "ERROR")
            messagebox.showerror("파싱 오류", f"추경 명세서 분석 중 오류가 발생했습니다:\n{str(e)}")

    def _on_add_rule_dialog(self):
        dialog = RuleEditDialog(self)
        self.wait_window(dialog)
        if dialog.result:
            self.rules.append(dialog.result)
            self._save_configs()
            self.classifier.set_rules(self.rules)
            self._render_rules_tab_data()
            self.log(f"➕ 새 수동 규칙 등록 완료: [{dialog.result['target_account']} > {dialog.result['target_sub_account']}]", "SUCCESS")
            self._reclassify_and_refresh()

    def _on_edit_rule_dialog(self):
        sel = self.tree_rules_view.selection()
        if not sel:
            messagebox.showwarning("선택 필요", "수정할 규칙을 목록에서 먼저 선택하세요.")
            return

        idx = self.tree_rules_view.index(sel[0])
        if 0 <= idx < len(self.rules):
            rule_data = self.rules[idx]
            dialog = RuleEditDialog(self, rule_data=rule_data)
            self.wait_window(dialog)
            if dialog.result:
                self.rules[idx] = dialog.result
                self._save_configs()
                self.classifier.set_rules(self.rules)
                self._render_rules_tab_data()
                self.log(f"✏ 규칙 수정 완료: [{dialog.result['target_account']} > {dialog.result['target_sub_account']}]", "SUCCESS")
                self._reclassify_and_refresh()

    def _on_delete_rule(self):
        sel = self.tree_rules_view.selection()
        if not sel:
            messagebox.showwarning("선택 필요", "삭제할 규칙을 목록에서 먼저 선택하세요.")
            return

        if not messagebox.askyesno("삭제 확인", "선택한 분류 규칙을 삭제하시겠습니까?"):
            return

        idx = self.tree_rules_view.index(sel[0])
        if 0 <= idx < len(self.rules):
            deleted = self.rules.pop(idx)
            self._save_configs()
            self.classifier.set_rules(self.rules)
            self._render_rules_tab_data()
            self.log(f"🗑 규칙 삭제 완료: [{deleted.get('target_sub_account')}]", "INFO")
            self._reclassify_and_refresh()

    def _on_export_rules(self):
        file_path = filedialog.asksaveasfilename(
            title="규칙 내보내기",
            defaultextension=".json",
            filetypes=[("JSON 파일 (*.json)", "*.json"), ("모든 파일", "*.*")],
            initialfile=f"예산관리대장_분류규칙_백업_{datetime.now().strftime('%Y%m%d')}.json"
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(self.rules, f, ensure_ascii=False, indent=2)
            self.log(f"💾 분류 규칙 파일 내보내기 완료: {file_path}", "SUCCESS")
            messagebox.showinfo("내보내기 완료", f"분류 규칙 {len(self.rules)}건을 성공적으로 저장했습니다.")

    def _on_import_rules(self):
        file_path = filedialog.askopenfilename(
            title="규칙 가져오기",
            filetypes=[("JSON 파일 (*.json)", "*.json"), ("모든 파일", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    imported = json.load(f)
                if isinstance(imported, list):
                    self.rules = imported
                    self._save_configs()
                    self.classifier.set_rules(self.rules)
                    self._render_rules_tab_data()
                    self.log(f"📂 외부 분류 규칙 {len(self.rules)}건 가져오기 완료: {file_path}", "SUCCESS")
                    self._reclassify_and_refresh()
                    messagebox.showinfo("가져오기 완료", f"규칙 {len(self.rules)}건을 성공적으로 불러왔습니다.")
            except Exception as e:
                messagebox.showerror("가져오기 오류", f"파일을 불러올 수 없습니다:\n{str(e)}")

    def _on_reset_rules(self):
        if messagebox.askyesno("기본값 초기화", "모든 규칙을 표준 행정 기본 규칙으로 초기화하시겠습니까?\n(직접 추가한 규칙은 삭제됩니다)"):
            self.rules = KeywordClassifier.get_default_rules()
            self._save_configs()
            self.classifier.set_rules(self.rules)
            self._render_rules_tab_data()
            self.log("🔄 분류 규칙을 행정 표준 기본 규칙으로 초기화했습니다.", "INFO")
            self._reclassify_and_refresh()

    def _on_double_click_tx(self, event=None):
        self._on_create_rule_from_selected_tx()

    def _on_create_rule_from_selected_tx(self):
        sel = self.tree_tx.selection()
        if not sel:
            messagebox.showwarning("선택 필요", "규칙을 생성할 지출 내역 행을 먼저 선택하세요.")
            return

        item_values = self.tree_tx.item(sel[0], "values")
        acc = item_values[2]
        summary = item_values[4]

        dialog = RuleEditDialog(self, initial_summary=summary)
        if acc:
            for idx, opt in enumerate(dialog.account_options):
                if acc in opt or opt in acc or acc.split()[0] == opt.split()[0]:
                    dialog.combo_account.current(idx)
                    break

        self.wait_window(dialog)
        if dialog.result:
            self.rules.append(dialog.result)
            self._save_configs()
            self.classifier.set_rules(self.rules)
            self._render_rules_tab_data()
            self.log(f"➕ 지출 건 기반 신규 수동 규칙 등록: [{dialog.result['target_sub_account']}]", "SUCCESS")
            self._reclassify_and_refresh()

    def _on_auto_infer_rules_manual(self):
        if not self.raw_transactions and not self.transactions:
            messagebox.showwarning("데이터 없음", "먼저 e호조 엑셀 파일을 분석하여 지출 내역을 불러오세요.")
            return

        target_txs = self.raw_transactions if self.raw_transactions else self.transactions
        all_rules, new_rules = KeywordClassifier.auto_infer_rules_from_transactions(target_txs, self.rules)
        if new_rules:
            self.rules = all_rules
            self._save_configs()
            self.classifier.set_rules(self.rules)
            self._render_rules_tab_data()
            self.log(f"🧠 신규 규칙 {len(new_rules)}건이 지출 적요 패턴에서 자동 추출되어 저장되었습니다.", "SUCCESS")
            self._reclassify_and_refresh()
            messagebox.showinfo("규칙 자동 추출 완료", f"{len(new_rules)}건의 신규 분류 규칙이 생성 및 등록되었습니다!")
        else:
            messagebox.showinfo("알림", "기존 규칙으로 모든 지출 패턴이 커버되고 있어 추가할 신규 규칙이 없습니다.")

    def _reclassify_and_refresh(self):
        source = self.raw_transactions if self.raw_transactions else self.transactions
        if not source:
            self.simulation_result = Forecaster.simulate(
                budget_items=self.budget_master,
                transactions=[],
                recurring_plans=self.recurring_plans,
                scheduled_plans=self.scheduled_plans,
                supplementary_items=self.supplementary_budgets
            )
            self._render_ledger_table()
            self._render_supp_table()
            self._render_monthly_table()
            return

        self.classifier.set_rules(self.rules)
        classified_txs = []
        classified_count = 0
        unclassified_count = 0

        for tx in source:
            summary = tx.get("summary", "")
            acc = tx.get("account", "")
            target_acc, sub_acc, rule_cond = self.classifier.classify(summary, acc)
            
            tx_copy = dict(tx)
            tx_copy["sub_account"] = sub_acc if sub_acc else "미분류"
            if target_acc and not acc:
                tx_copy["account"] = target_acc
            tx_copy["rule_matched"] = rule_cond if rule_cond else ""

            if tx_copy["sub_account"] != "미분류":
                classified_count += 1
            else:
                unclassified_count += 1

            classified_txs.append(tx_copy)

        self.transactions = classified_txs

        self.simulation_result = Forecaster.simulate(
            budget_items=self.budget_master,
            transactions=self.transactions,
            recurring_plans=self.recurring_plans,
            scheduled_plans=self.scheduled_plans,
            supplementary_items=self.supplementary_budgets
        )

        try:
            output_file = os.path.join(self.output_dir, f"{datetime.now().year}년_세출예산_세목별_예산관리대장.xlsx")
            ExcelExporter.export_report(self.simulation_result, self.transactions, output_file)
            self.last_output_file = output_file
        except Exception:
            pass

        self._render_ledger_table()
        self._render_supp_table()
        self._render_monthly_table()
        self._render_transactions_table()
        self.log(f"🔄 실시간 재분류 완료: 분류 성공 {classified_count}건 / 미분류 {unclassified_count}건 (예산대장·추경·월별통계 자동 갱신됨)", "INFO")

    def _on_start_analysis(self):
        file_path = self.selected_file_path.get().strip()
        if not file_path:
            sample_path = os.path.join(self.base_dir, "sample", "e호조_23310_샘플.xlsx")
            if not os.path.exists(sample_path):
                from sample_generator import generate_sample_file, generate_sample_base_budget_txt, generate_sample_supplementary_txt
                generate_sample_file(sample_path)
                generate_sample_base_budget_txt(os.path.join(self.base_dir, "sample", "2026년도_본예산_세출예산명세서_샘플.txt"))
                generate_sample_supplementary_txt(os.path.join(self.base_dir, "sample", "제1회_추경_세출사업명세서_샘플.txt"))
            self.selected_file_path.set(sample_path)
            file_path = sample_path
            self.log("지정된 파일이 없어 테스트용 샘플 엑셀 파일로 분석을 시작합니다.", "WARN")

        if not os.path.exists(file_path):
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{file_path}")
            return

        self.log(f"3단계 e호조 지출 분석 시작: {file_path}", "INFO")

        try:
            raw_txs, parse_msg = EHojoParser.parse_file(file_path)
            self.raw_transactions = raw_txs
            self.log(parse_msg, "SUCCESS")
        except Exception as e:
            self.log(f"파일 파싱 오류: {str(e)}", "ERROR")
            messagebox.showerror("파싱 오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")
            return

        all_rules, new_rules = KeywordClassifier.auto_infer_rules_from_transactions(raw_txs, self.rules)
        if new_rules:
            self.rules = all_rules
            self._save_configs()
            self.classifier.set_rules(self.rules)
            self._render_rules_tab_data()
            self.log(f"🧠 업로드 지출 패턴 분석: 신규 분류 규칙 {len(new_rules)}건 자동 추출 및 등록 완료!", "SUCCESS")

        self._reclassify_and_refresh()

        classified_count = sum(1 for tx in self.transactions if tx.get("sub_account") != "미분류")
        unclassified_count = len(self.transactions) - classified_count

        rule_msg = f"\n• ✨ 신규 자동 생성된 규칙: {len(new_rules)}건" if new_rules else ""
        messagebox.showinfo(
            "분석 및 자동규칙 적용 완료",
            f"분석, 12.31 연말 예측 및 월별 통계 집계가 완료되었습니다!\n\n"
            f"• 총 지출건수: {len(self.transactions)}건\n"
            f"• 자동분류 성공: {classified_count}건 (미분류: {unclassified_count}건){rule_msg}\n"
            f"• 12.31 예상 집행총액: {self.simulation_result.get('total_forecast_spent', 0):,}원\n"
            f"• 12.31 예상 불용잔액: {self.simulation_result.get('total_forecast_balance', 0):,}원\n\n"
            f"결과 엑셀 파일이 'output' 폴더에 생성되었습니다."
        )

    def _render_ledger_table(self):
        for item in self.tree_ledger.get_children():
            self.tree_ledger.delete(item)

        sim = self.simulation_result
        if not sim:
            return

        self.lbl_ledger_summary.config(
            text=f"총 배정예산(추경후): {sim['total_budget']:,}원  |  "
                 f"기집행액: {sim['total_spent']:,}원 (집행률: {sim['overall_exec_rate']}%)  |  "
                 f"현재잔액: {sim['current_balance']:,}원  |  "
                 f"12.31 예상잔액: {sim['total_forecast_balance']:,}원 (최종예상: {sim['overall_forecast_exec_rate']}%)"
        )

        current_acc = None
        for it in sim.get("items", []):
            acc = it["account"]
            if acc != current_acc:
                current_acc = acc
                self.tree_ledger.insert("", tk.END, values=(
                    it.get("detail_project", "-"),
                    it.get("category", "-"),
                    acc, "──────────────", "", "", "", "", "", "", "", "", ""
                ), tags=("header_row",))

            tag = ()
            if "초과" in it["status"]:
                tag = ("danger_row",)
            elif "불용" in it["status"]:
                tag = ("warning_row",)

            self.tree_ledger.insert("", tk.END, values=(
                it.get("detail_project", "-"),
                it.get("category", "-"),
                it["account"],
                it["sub_account"],
                f"{it['budget']:,}",
                f"{it['actual_spent']:,}",
                f"{it['current_balance']:,}",
                f"{it['exec_rate']}%",
                f"{it['remaining_recurring']:,}",
                f"{it['scheduled_spent']:,}",
                f"{it['forecast_total_spent']:,}",
                f"{it['forecast_balance']:,}",
                it["status"]
            ), tags=tag)

        self.tree_ledger.insert("", tk.END, values=(
            "합 계", "전체", "총괄", "전체 세목",
            f"{sim['total_budget']:,}",
            f"{sim['total_spent']:,}",
            f"{sim['current_balance']:,}",
            f"{sim['overall_exec_rate']}%",
            "-",
            "-",
            f"{sim['total_forecast_spent']:,}",
            f"{sim['total_forecast_balance']:,}",
            f"최종 {sim['overall_forecast_exec_rate']}%"
        ), tags=("total_row",))

    def _render_supp_table(self):
        for item in self.tree_supp.get_children():
            self.tree_supp.delete(item)

        supp_data = self.simulation_result.get("supplementary_matrix", {})
        if not supp_data or not supp_data.get("rows"):
            return

        s_rows = supp_data.get("rows", [])
        current_acc = None

        for sr in s_rows:
            acc = sr["account"]
            if acc != current_acc:
                current_acc = acc
                self.tree_supp.insert("", tk.END, values=(
                    sr.get("detail_project", "-"),
                    acc, "──────────────", "", "", "", "", "", "", "", "", "", ""
                ), tags=("header_row",))

            def fmt_diff(v):
                if v > 0:
                    return f"+{v:,}"
                elif v < 0:
                    return f"{v:,}"
                return "-"

            self.tree_supp.insert("", tk.END, values=(
                sr.get("detail_project", "-"),
                sr["account"],
                sr["sub_account"],
                f"{sr['base_budget']:,}",
                fmt_diff(sr["r1"]),
                fmt_diff(sr["r2"]),
                fmt_diff(sr["r3"]),
                fmt_diff(sr["r4"]),
                f"{sr['final_budget']:,}",
                f"{sr['spent']:,}",
                f"{sr['balance']:,}",
                f"{sr['exec_rate']}%",
                sr["reason"]
            ))

        def fmt_diff_tot(v):
            if v > 0:
                return f"+{v:,}"
            elif v < 0:
                return f"{v:,}"
            return "-"

        self.tree_supp.insert("", tk.END, values=(
            "합 계", "전체", "전체 추경 총괄",
            f"{supp_data.get('total_base', 0):,}",
            fmt_diff_tot(supp_data.get("total_r1", 0)),
            fmt_diff_tot(supp_data.get("total_r2", 0)),
            fmt_diff_tot(supp_data.get("total_r3", 0)),
            fmt_diff_tot(supp_data.get("total_r4", 0)),
            f"{supp_data.get('total_final', 0):,}",
            f"{supp_data.get('total_spent', 0):,}",
            f"{supp_data.get('total_balance', 0):,}",
            f"{supp_data.get('overall_exec_rate', 0)}%",
            "-"
        ), tags=("total_row",))

    def _render_monthly_table(self):
        for item in self.tree_monthly.get_children():
            self.tree_monthly.delete(item)

        m_data = self.simulation_result.get("monthly_matrix", {})
        if not m_data:
            return

        m_rows = m_data.get("rows", [])
        current_acc = None

        for r in m_rows:
            acc = r["account"]
            if acc != current_acc:
                current_acc = acc
                self.tree_monthly.insert("", tk.END, values=(
                    r.get("detail_project", "-"),
                    acc, "──────────────", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
                ), tags=("header_row",))

            m_vals = [f"{v:,}" if v > 0 else "-" for v in r["months"]]

            self.tree_monthly.insert("", tk.END, values=(
                r.get("detail_project", "-"),
                r["account"],
                r["sub_account"],
                f"{r['budget']:,}",
                *m_vals,
                f"{r['total_spent']:,}",
                f"{r['balance']:,}",
                f"{r['exec_rate']}%"
            ))

        m_tot_vals = [f"{v:,}" if v > 0 else "-" for v in m_data.get("monthly_totals", [0] * 12)]
        self.tree_monthly.insert("", tk.END, values=(
            "합 계", "전체", "전체 월별 총괄",
            f"{m_data.get('total_budget', 0):,}",
            *m_tot_vals,
            f"{m_data.get('total_spent', 0):,}",
            f"{m_data.get('total_balance', 0):,}",
            f"{m_data.get('overall_exec_rate', 0)}%"
        ), tags=("total_row",))

    def _render_transactions_table(self):
        for item in self.tree_tx.get_children():
            self.tree_tx.delete(item)

        self.lbl_tx_count.config(text=f"총 {len(self.transactions)}건의 지출 내역")

        for idx, tx in enumerate(self.transactions, start=1):
            sub_acc = tx.get("sub_account", "미분류")
            tags = () if sub_acc != "미분류" else ("unclassified",)
            self.tree_tx.insert("", tk.END, values=(
                idx,
                tx.get("date", ""),
                tx.get("account", ""),
                sub_acc,
                tx.get("summary", ""),
                tx.get("vendor", ""),
                f"{int(tx.get('amount', 0)):,}",
                tx.get("rule_matched", "")
            ), tags=tags)

    def _filter_transactions(self, event=None):
        q = self.entry_search_tx.get().strip().lower()
        for item in self.tree_tx.get_children():
            self.tree_tx.delete(item)

        filtered = []
        for idx, tx in enumerate(self.transactions, start=1):
            text_pool = f"{tx.get('summary', '')} {tx.get('vendor', '')} {tx.get('sub_account', '')} {tx.get('account', '')}".lower()
            if not q or q in text_pool:
                filtered.append((idx, tx))

        self.lbl_tx_count.config(text=f"검색 결과: {len(filtered)} / 총 {len(self.transactions)}건")
        for idx, tx in filtered:
            sub_acc = tx.get("sub_account", "미분류")
            tags = () if sub_acc != "미분류" else ("unclassified",)
            self.tree_tx.insert("", tk.END, values=(
                idx,
                tx.get("date", ""),
                tx.get("account", ""),
                sub_acc,
                tx.get("summary", ""),
                tx.get("vendor", ""),
                f"{int(tx.get('amount', 0)):,}",
                tx.get("rule_matched", "")
            ), tags=tags)

    def _on_open_excel(self):
        if not os.path.exists(self.last_output_file):
            messagebox.showwarning("파일 없음", "먼저 '업로드 파일 분석 시작'을 실행하여 엑셀 파일을 생성하세요.")
            return
        try:
            os.startfile(self.last_output_file)
            self.log(f"엑셀 파일 열기 실행: {self.last_output_file}", "INFO")
        except Exception as e:
            messagebox.showerror("열기 오류", f"엑셀 파일을 열 수 없습니다:\n{str(e)}")

    def _on_open_folder(self):
        try:
            os.startfile(self.output_dir)
            self.log(f"결과 저장 폴더 열기: {self.output_dir}", "INFO")
        except Exception as e:
            messagebox.showerror("열기 오류", f"폴더를 열 수 없습니다:\n{str(e)}")


if __name__ == "__main__":
    app = BudgetApp()
    app.mainloop()
