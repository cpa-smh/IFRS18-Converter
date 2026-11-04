from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
import importlib.util
import pytest
from openpyxl import Workbook,load_workbook
from ifrs18.models import Item,Company
from ifrs18.engine import classify
from ifrs18.ai import recommend,Recommendation
from ifrs18.parser import parse,detect
from convert import convert


def item(name='外환',**kwargs):
    return Item('1',name,Decimal(10),Decimal(8),'A2',**kwargs)


def test_rules_and_business_exceptions():
    ordinary=Company('제조',2025,activities='제품 제조')
    bank=Company('은행',2025,'금융업','고객 대출',finances_customers=True)
    assert classify(item('매출액'),ordinary).category=='영업'
    assert classify(item('지분법이익'),ordinary).category=='투자'
    assert classify(item('처분이익',source_type='operating_ppe',note='제조용 설비'),ordinary).category=='영업'
    assert classify(item('이자비용',source_type='customer_funding',note='고객대출 자금 조달'),bank).category=='영업'
    assert classify(item('차입이자',source_type='loan_funding',note='부채'),bank).category=='판단 보류'


def test_missing_information_hold():
    assert classify(item('외환차익'),Company('제조',2025,activities='제품 제조')).category=='판단 보류'
    result,error=recommend(item(),Company('제조',2025),api_key=None)
    assert result is None and '미연결' in error


def test_ai_structured_recommendation_and_no_approval():
    i=item('외환차익',note='제품 매출채권에서 발생한 환차익')
    data=Recommendation(account_name=i.name,original_amount='10',category='영업',reason='매출채권 원천',evidence=[{'source_id':'note','quote':'매출채권'}],paragraphs=['B65'],confidence=.9,needs_review=False,required_information=['매출채권 외환차익 명세'])
    client=Mock();client.responses.parse.return_value=SimpleNamespace(output_parsed=data)
    result,error=recommend(i,Company('제조',2025),client=client)
    assert not error and result['needs_review'] is True and result['required_information']
    assert client.responses.parse.call_args.kwargs['store'] is False


def test_ai_invalid_evidence_rejected():
    data=Recommendation(account_name='외환차익',original_amount='10',category='재무',reason='부채',evidence=[{'source_id':'note','quote':'차입금'}],paragraphs=['B65'],confidence=.9,needs_review=True,required_information=[])
    client=Mock();client.responses.parse.return_value=SimpleNamespace(output_parsed=data)
    result,error=recommend(item('외환차익',note='매출채권'),Company('제조',2025),client=client)
    assert result is None and '인용 검증' in error


def test_ai_failure_safe():
    client=Mock();client.responses.parse.side_effect=RuntimeError('secret_key_123')
    result,error=recommend(item(note='명세'),Company('제조',2025),client=client)
    assert result is None and 'secret' not in error


def test_excel_preserves_source_and_has_live_formulas(tmp_path):
    w=Workbook();s=w.active;s.title='원본';s.append(['계정과목','당기','전기','기존범주','주석']);s.append(['매출액',100,90,'영업','']);s.append(['매출원가',-40,-30,'영업','']);s.append(['당기순손익',60,60,'','']);s['F1']='=B2+B3';s.merge_cells('H1:I1')
    p=tmp_path/'in.xlsx';w.save(p);before=p.read_bytes()
    result=convert(p,tmp_path/'out.xlsx',Company('제조',2025,activities='제품 제조'),use_ai=False)
    out=load_workbook(tmp_path/'out.xlsx')
    assert p.read_bytes()==before and out['원본']['F1'].value=='=B2+B3'
    assert 'H1:I1' in out['원본'].merged_cells
    assert out['공시IS_IFRS18']['B15'].value=='=B11+SUM(B13:B14)'
    assert out['공시IS_IFRS18']['B24'].value=='=B22+B23'
    assert len(out['계정재분류'].data_validations.dataValidation)==2
    assert result['accounts']==2
    assert out['계정재분류']['T2'].data_type=='f'
    with pytest.raises(ValueError):convert(p,p,Company('제조',2025))


def test_missing_formula_cache_flagged(tmp_path):
    w=Workbook();s=w.active;s.append(['계정과목','당기']);s.append(['매출액','=10+2']);p=tmp_path/'f.xlsx';w.save(p)
    items,_=parse(p,detect(p)[0]);assert items[0].issues and items[0].current is None


def test_ambiguous_layout_requires_mapping(tmp_path):
    w=Workbook();s=w.active;s.append(['계정과목','당기','당기']);s.append(['매출액',1,2]);p=tmp_path/'a.xlsx';w.save(p)
    assert detect(p)==[]
