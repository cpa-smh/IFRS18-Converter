from decimal import Decimal
from types import SimpleNamespace
from unittest.mock import Mock,patch
from openpyxl import load_workbook
from ifrs18.models import Item,Company
from ifrs18.dart_notes import sections,attach,business_context,locations
from ifrs18.batch_ai import classify_batch,Batch,Row
from examples.make_example import make
from convert import convert


def corpus():
    return {'receipt':'20260312001224','basis':'CFS','source_url':'https://dart.fss.or.kr/dsaf001/main.do?rcpNo=20260312001224',
        'documents':[{'file':'audit.xml','role':'audit','text':'1. 일반사항\n자동차 제조\n2. 금융손익\n'+('긴 설명 '*1800)+'\n이자비용은 차입금에서 발생 (주석 3 참조)\n1 | 123 | 표 안 숫자\n3. 차입금\n은행 차입 이자\n4. 유형자산\n설비 처분\n1. 외부감사 실시내용\n주석이 아님'},
            {'file':'main.xml','role':'business_report','text':'I. 회사의 개요\n개요\nII. 사업의 내용\n1. 사업의 개요\n자동차 제조 및 판매\n2. 주요 제품\n제품 정보\nIII. 재무에 관한 사항\n별도 ONLY'}]}


def item():return Item('x','금융비용',Decimal(-10),Decimal(-9),'A1')


def test_complete_notes_and_cross_reference():
    c=corpus();i=item();attach([i],c)
    assert len(i.note)>6000 and '이자비용은 차입금' in i.note and '은행 차입 이자' in i.note
    assert '유형자산' not in i.note and '외부감사 실시내용' not in i.note
    assert [s['number'] for s in sections(c) if s['role']=='audit']==['1','2','3','4']


def test_business_boundaries_and_locations():
    c=corpus();i=item();attach([i],c);b=business_context(c)
    assert '자동차 제조 및 판매' in b and '제품 정보' in b and '별도 ONLY' not in b
    loc=locations(i,[{'source_id':'note','quote':'은행 차입 이자'}],b)
    assert '연결 주석 3 차입금' in loc[0] and 'rcpNo=' in loc[0]


def row(category='판단 보류',quote='자동차 제조',requests=None):
    return Row(item_id='x',account_name='금융비용',original_amount='-10',category=category,reason='차입금에서 발생',
        evidence=[{'source_id':'note','quote':quote}],paragraphs=['60'],confidence=.9,needs_review=False,
        required_information=[] if category!='판단 보류' else ['차입금 주석'],requested_sections=requests or [])


def test_request_additional_note_and_apply():
    c=corpus();i=item();attach([i],c)
    # 최초 연결에 없는 유형자산 주석을 요청하는 과정을 모의 실행한다.
    target=next(s for s in sections(c) if s['number']=='4')
    client=Mock();client.responses.parse.side_effect=[SimpleNamespace(output_parsed=Batch(results=[row(requests=[target['id']])])),
        SimpleNamespace(output_parsed=Batch(results=[row('영업','설비 처분')]))]
    result=classify_batch([i],Company('기아',2025),None,'model',client,corpus=c)
    assert client.responses.parse.call_count==2
    assert result['x'][0]['category']=='영업' and '주석 4' in result['x'][0]['source_locations'][0]


def test_invalid_quote_rejected():
    c=corpus();i=item();attach([i],c);client=Mock()
    client.responses.parse.return_value=SimpleNamespace(output_parsed=Batch(results=[row('재무','없는 문구')]))
    assert classify_batch([i],Company('기아',2025),None,'model',client,corpus=c)['x'][0] is None


def test_api_error_is_not_evidence_shortage():
    client=Mock();client.responses.parse.side_effect=RuntimeError('secret')
    result=classify_batch([item()],Company('기아',2025),None,'model',client)
    assert '아직 판단하지 못함' in result['x'][1] and 'secret' not in result['x'][1]


def test_excel_source_location_and_complete_source(tmp_path):
    src=make(tmp_path);c=corpus()
    ai={'account_name':'외환차익','original_amount':'1','category':'영업','reason':'매출채권 원천 확인',
        'evidence':[{'source_id':'note','quote':'은행 차입 이자'}],'paragraphs':['B65'],'confidence':.9,
        'needs_review':False,'required_information':[],'source_locations':['연결 주석 3 차입금']}
    from ifrs18.parser import parse,detect
    items,_=parse(src,detect(src)[0]);fx=next(i for i in items if i.name=='외환차익')
    # 공시 회사명 검증에 맞춘 이름이며, 실제 AI 호출은 하지 않는다.
    with patch('convert.collect',return_value=c),patch('convert.classify_batch',return_value={fx.id:(ai,'')}):
        convert(src,tmp_path/'out.xlsx',Company('자동차',2025),receipt=c['receipt'])
    w=load_workbook(tmp_path/'out.xlsx');s=w['계정재분류']
    r=next(r for r in range(2,s.max_row+1) if s.cell(r,2).value=='외환차익')
    assert s.cell(r,21).value=='연결 주석 3 차입금' and s.cell(r,14).value=='자동 적용'
    assert '공시원문근거' in w.sheetnames
    assert any('긴 설명' in str(row[4].value) for row in w['공시원문근거'].iter_rows(min_row=2))
