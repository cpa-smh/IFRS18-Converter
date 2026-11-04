from dataclasses import dataclass, field, asdict
from decimal import Decimal, InvalidOperation
import hashlib
import json

CATEGORIES=('영업','투자','재무','법인세','중단영업')
CHOICES=(*CATEGORIES,'복수 범주 분할','판단 보류')


def money(value):
    if value is None or isinstance(value,bool): return None
    text=str(value).strip().replace(',','').replace('−','-')
    if text in ('','-','—'): return None
    if text.startswith('(') and text.endswith(')'): text='-'+text[1:-1]
    try:
        result=Decimal(text)
        return result if result.is_finite() else None
    except InvalidOperation: return None


@dataclass
class Company:
    name: str
    year: int
    industry: str='제조업'
    activities: str=''
    invests_assets: bool=False
    invests_financial_assets: bool=False
    finances_customers: bool=False
    unrelated_policy: str='투자·재무'  # 56(2)(b)/65(1)(b)의 일관된 선택


@dataclass
class Item:
    id: str
    name: str
    current: Decimal | None
    previous: Decimal | None
    source: str
    old_category: str=''
    description: str=''
    note: str=''
    source_type: str='unknown'
    role: str='detail'
    issues: list[str]=field(default_factory=list)


@dataclass
class Decision:
    category: str='판단 보류'
    method: str='보류'
    reason: str='원천과 경제적 실질 확인 필요'
    refs: list[str]=field(default_factory=list)
    group: str='other'
    ai: dict | None=None


def fingerprint(items):
    return hashlib.sha256(json.dumps([asdict(i) for i in items],ensure_ascii=False,sort_keys=True,default=str).encode()).hexdigest()
