"""
키워드 기반 자동 분류 엔진 및 업로드 내역 기반 규칙 자동 생성/학습 엔진
(Keyword Classifier & Auto Rule Inferrer)
- '키워드1 AND 키워드2', '키워드1 OR 키워드2', '키워드1 NOT 키워드2' 문법 지원
- 수동 규칙(priority=20)이 자동 생성 규칙(priority=10)보다 우선 적용
- 규칙 백업(내보내기/가져오기) 및 기본값 초기화 지원
"""

import json
import re
from typing import List, Dict, Any, Optional, Tuple, Set


class KeywordClassifier:
    ADMIN_KNOWLEDGE_BASE = [
        {
            "account_prefix": "201-01",
            "account_name": "201-01 사무관리비",
            "sub_account": "복사용지 및 사무용품비",
            "keywords": ["복사용지", "A4", "토너", "잉크", "사무용품", "문구", "화일", "바인더", "필기구", "스테이플러", "드림디포", "오피스", "소모품"],
            "not_keywords": ["현수막", "배너", "간담회"]
        },
        {
            "account_prefix": "201-01",
            "account_name": "201-01 사무관리비",
            "sub_account": "간담회 다과비",
            "keywords": ["다과", "음료", "간담회", "차류", "다과용품", "회의용", "베이커리", "커피", "떡", "샌드위치", "케이크", "파리바게뜨", "뚜레쥬르"],
            "not_keywords": ["식대", "급량비", "야근"]
        },
        {
            "account_prefix": "201-01",
            "account_name": "201-01 사무관리비",
            "sub_account": "신문 및 정기간행물 구독료",
            "keywords": ["신문", "일간지", "정기간행물", "구독료", "신문대금", "잡지", "보급소", "지방지", "매일신문", "조선", "동아", "중앙", "부산일보", "국제신문"],
            "not_keywords": []
        },
        {
            "account_prefix": "201-01",
            "account_name": "201-01 사무관리비",
            "sub_account": "직원 급량비",
            "keywords": ["급량비", "매식비", "야근식대", "특근", "비상근무", "식대", "도시락", "야근", "식당", "김밥"],
            "not_keywords": ["다과", "음료"]
        },
        {
            "account_prefix": "201-01",
            "account_name": "201-01 사무관리비",
            "sub_account": "현수막 및 홍보인쇄물",
            "keywords": ["현수막", "배너", "인쇄물", "리플릿", "책자", "포스터", "홍보물", "전단", "인쇄", "디자인", "간판", "배너제작", "카탈로그"],
            "not_keywords": []
        },
        {
            "account_prefix": "201-01",
            "account_name": "201-01 사무관리비",
            "sub_account": "피복비",
            "keywords": ["피복비", "근무복", "안전화", "방한복", "조끼", "작업복", "모자", "장화", "피복", "유니폼"],
            "not_keywords": []
        },
        {
            "account_prefix": "201-01",
            "account_name": "201-01 사무관리비",
            "sub_account": "우편요금 및 발송비",
            "keywords": ["우편", "등기", "택배", "발송", "우체국", "내용증명", "소포", "등기우편"],
            "not_keywords": []
        },
        {
            "account_prefix": "201-01",
            "account_name": "201-01 사무관리비",
            "sub_account": "도서구입비",
            "keywords": ["도서", "서적", "도서구입", "출판사", "문고", "교보문고", "영풍문고"],
            "not_keywords": ["신문", "구독"]
        },
        {
            "account_prefix": "201-02",
            "account_name": "201-02 공공운영비",
            "sub_account": "청사 전기 및 가스요금",
            "keywords": ["전기요금", "한전", "전기료", "도시가스", "가스요금", "난방비", "한국전력", "가스", "열요금", "전력"],
            "not_keywords": []
        },
        {
            "account_prefix": "201-02",
            "account_name": "201-02 공공운영비",
            "sub_account": "통신 및 전산회선료",
            "keywords": ["통신요금", "회선료", "전용회선", "인터넷", "전화요금", "KT", "SKT", "LGU", "행정전화", "전화료", "전산회선"],
            "not_keywords": []
        },
        {
            "account_prefix": "201-02",
            "account_name": "201-02 공공운영비",
            "sub_account": "정수기 및 복사기 렌탈료",
            "keywords": ["정수기", "렌탈료", "임차료", "복사기임차", "코웨이", "청호나이스", "SK매직", "복합기", "렌탈", "필터교체"],
            "not_keywords": []
        },
        {
            "account_prefix": "201-02",
            "account_name": "201-02 공공운영비",
            "sub_account": "공용차량 유지관리비",
            "keywords": ["차량", "유류비", "주유대", "엔진오일", "정비", "차량보험", "세차", "주유", "GS칼텍스", "SK에너지", "블루핸즈", "오토오아시스", "타이어"],
            "not_keywords": []
        },
        {
            "account_prefix": "201-02",
            "account_name": "201-02 공공운영비",
            "sub_account": "청사 환경관리 및 용역비",
            "keywords": ["청소", "방역", "소독", "정화조", "환경관리", "용역비", "청소용역", "경비용역"],
            "not_keywords": []
        },
        {
            "account_prefix": "202-01",
            "account_name": "202-01 국내여비",
            "sub_account": "시정업무추진 관내여비",
            "keywords": ["관내", "시내", "관내출장", "출장여비", "출장", "순찰", "점검"],
            "not_keywords": ["세종", "서울", "국회", "관외", "KTX", "대전"]
        },
        {
            "account_prefix": "202-01",
            "account_name": "202-01 국내여비",
            "sub_account": "중앙부처 방문 관외여비",
            "keywords": ["관외", "세종", "서울", "대전", "국회", "KTX", "관외출장", "중앙부처", "정부청사", "항공", "고속버스"],
            "not_keywords": []
        },
        {
            "account_prefix": "203-04",
            "account_name": "203-04 부서운영업무추진비",
            "sub_account": "부서운영 간담회 및 격려비",
            "keywords": ["부서운영", "업무추진비", "부서간담회", "직원격려", "시책추진", "격려비", "소통간담회", "간담회"],
            "not_keywords": []
        }
    ]

    def __init__(self, rules: Optional[List[Dict[str, Any]]] = None):
        self.rules = rules or []

    def set_rules(self, rules: List[Dict[str, Any]]):
        # 우선순위 내림차순 정렬 (수동 규칙 20 > 자동생성 규칙 10)
        self.rules = sorted(rules, key=lambda x: x.get("priority", 10), reverse=True)

    @classmethod
    def evaluate_condition(cls, condition_str: str, text: str) -> bool:
        if not condition_str or not condition_str.strip():
            return False
        
        target_text = text.lower()
        expr = condition_str.strip()

        tokens = re.findall(r'\(|\)|\bAND\b|\bOR\b|\bNOT\b|[^\s\(\)]+', expr, flags=re.IGNORECASE)
        
        py_expr_parts = []
        for t in tokens:
            t_upper = t.upper()
            if t_upper in ("AND", "OR", "NOT", "(", ")"):
                py_expr_parts.append(t_upper.lower() if t_upper != "(" and t_upper != ")" else t_upper)
            else:
                matched = (t.lower() in target_text)
                py_expr_parts.append(str(matched))

        py_expr = " ".join(py_expr_parts)
        
        try:
            result = eval(py_expr, {"__builtins__": {}}, {})
            return bool(result)
        except Exception:
            cleaned = re.sub(r'[\(\)]|\bAND\b|\bOR\b|\bNOT\b', ' ', expr, flags=re.IGNORECASE).split()
            return any(k.lower() in target_text for k in cleaned if k.strip())

    def classify(self, text: str, account_code: Optional[str] = None) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        if not text:
            return None, "미분류", None

        for rule in self.rules:
            rule_account = rule.get("target_account", "")
            if account_code and rule_account:
                if account_code.split()[0] != rule_account.split()[0] and account_code != rule_account:
                    continue

            condition = rule.get("condition", "")
            if self.evaluate_condition(condition, text):
                return rule.get("target_account"), rule.get("target_sub_account"), condition

        return account_code, "미분류", None

    @classmethod
    def auto_infer_rules_from_transactions(
        cls,
        transactions: List[Dict[str, Any]],
        existing_rules: List[Dict[str, Any]]
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """
        e호조 지출 엑셀 내역(적요 및 채권자)을 분석하여 매칭 가능한 자동분류 규칙을 자동으로 추출/생성합니다.
        - 수동 규칙(priority=20)은 절대 덮어쓰지 않고 보존
        - 신규 자동 생성 규칙은 priority=10 부여
        """
        new_rules = []
        # (target_account, target_sub_account) -> rule
        rule_map = {}
        for r in existing_rules:
            key = (r.get("target_account", "").strip(), r.get("target_sub_account", "").strip())
            rule_map[key] = r

        for tx in transactions:
            acc = tx.get("account", "").strip()
            summary = tx.get("summary", "").strip()
            vendor = tx.get("vendor", "").strip()
            combined_text = f"{summary} {vendor}".lower()

            if not combined_text:
                continue

            for kb_item in cls.ADMIN_KNOWLEDGE_BASE:
                kb_acc_prefix = kb_item["account_prefix"]
                kb_acc_name = kb_item["account_name"]
                kb_sub_acc = kb_item["sub_account"]

                if acc and not acc.startswith(kb_acc_prefix) and kb_acc_name != acc:
                    continue

                matched_keywords = [k for k in kb_item["keywords"] if k.lower() in combined_text]
                if matched_keywords:
                    key = (kb_acc_name, kb_sub_acc)
                    
                    if key not in rule_map:
                        condition_parts = [k for k in kb_item["keywords"][:6]]
                        cond_str = " OR ".join(condition_parts)
                        if kb_item["not_keywords"]:
                            not_str = " OR ".join(kb_item["not_keywords"][:3])
                            cond_str = f"({cond_str}) AND NOT ({not_str})"

                        new_rule = {
                            "target_account": kb_acc_name,
                            "target_sub_account": kb_sub_acc,
                            "condition": cond_str,
                            "priority": 10, # 자동 생성은 우선순위 10
                            "auto_generated": True
                        }
                        new_rules.append(new_rule)
                        rule_map[key] = new_rule

        all_rules = sorted(list(rule_map.values()), key=lambda x: x.get("priority", 10), reverse=True)
        return all_rules, new_rules

    @classmethod
    def get_default_rules(cls) -> List[Dict[str, Any]]:
        """기본 표준 규칙 목록 반환"""
        defaults = []
        for kb in cls.ADMIN_KNOWLEDGE_BASE:
            cond = " OR ".join(kb["keywords"][:6])
            if kb["not_keywords"]:
                cond = f"({cond}) AND NOT ({' OR '.join(kb['not_keywords'][:3])})"
            defaults.append({
                "target_account": kb["account_name"],
                "target_sub_account": kb["sub_account"],
                "condition": cond,
                "priority": 10,
                "auto_generated": False
            })
        return defaults
