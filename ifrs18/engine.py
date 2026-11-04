from pathlib import Path
import json
import re
from .models import Decision

RULES=json.loads(Path(__file__).with_name('rules.json').read_text(encoding='utf-8'))
def normalize(s): return re.sub(r'\s+','',s)


def classify(item,company):
    if item.role=='oci': return Decision('OCI','규칙','기타포괄손익은 당기손익 5개 범주와 별도 표시',['89'],'oci')
    if item.role!='detail': return Decision('제외','통제','원본 중간합계/당기순손익은 상세 계정에 중복 합산하지 않음')
    if item.issues: return Decision(reason='; '.join(item.issues))
    if not company.activities.strip(): return Decision(reason='회사의 주된 사업활동 입력 필요')
    source=RULES['source_rules'].get(item.source_type)
    if source and not (item.note.strip() or item.description.strip()):
        return Decision(reason='발생 원천 코드의 근거 설명/주석 필요')
    if source:
        category=source['category']; refs=source['refs']
        if item.source_type in ('loan_funding','loan_fx') and company.finances_customers:
            return Decision(reason='고객 금융제공 기업: 차입의 관련성·65항 회계정책 확인 필요',refs=['65','66'])
        if item.source_type=='customer_funding':
            category='영업' if company.finances_customers else '재무'
        if item.source_type=='independent_asset' and company.invests_assets: category='영업'
        if item.source_type=='cash_equivalent':
            if company.invests_financial_assets: category='영업'
            elif company.finances_customers:
                return Decision(reason='금융제공 관련 현금 여부와 문단 56~57의 정책 선택 확인 필요',refs=['56','57'])
        return Decision(category,'규칙',source['reason'],refs,source['group'])
    # 보험·금융업의 명칭만으로는 주된 사업활동 판단을 완료하지 않는다.
    if any(x in company.industry for x in ('금융','은행','보험')):
        return Decision(reason='금융·보험업: 구체적인 수익·부채의 경제적 원천 확인 필요',refs=['49','65'])
    matches=[r for r in RULES['rules'] if normalize(item.name) in [normalize(n) for n in r['names']]]
    if len(matches)==1:
        r=matches[0];return Decision(r['category'],'규칙',r['reason'],r['refs'],r['group'])
    return Decision(reason='명칭만으로 발생 원천을 확정할 수 없음. 주석·명세 또는 사용자 판단 필요',refs=['47','B65'] if any(t in item.name for t in ('외환','환산')) else ['47','49'])
