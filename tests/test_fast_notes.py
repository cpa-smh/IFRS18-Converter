from io import BytesIO
from zipfile import ZipFile
from types import SimpleNamespace
from unittest.mock import Mock,patch
from decimal import Decimal
from openpyxl import load_workbook
from ifrs18.models import Item,Company
from ifrs18.dart_notes import parse_zip,attach,collect
from ifrs18.batch_ai import classify_batch,Batch,Row
from examples.make_example import make
from convert import convert


def test_document_scope_and_cache(tmp_path):
    b=BytesIO()
    with ZipFile(b,'w') as z:
        z.writestr('20260310002820.xml','<r>사업의 내용: 제품 제조</r>')
        z.writestr('a_00761.xml','<r><p>금융수익 및 금융비용</p><table><tr><td>이자비용</td><td>100</td></tr></table></r>')
        z.writestr('a_00760.xml','<r>별도 ONLY</r>')
    d=parse_zip(b.getvalue(),'20260310002820','CFS')
    assert '별도 ONLY' not in str(d) and len(d['documents'])==2
    i=Item('x','금융비용',Decimal(-100),Decimal(-90),'A1')
    attach([i],d);assert '이자비용' in i.note and 'rcpNo=' in i.note
    import json
    p=tmp_path/'20260310002820_CFS.json';p.write_text(json.dumps(d),encoding='utf-8')
    assert collect('20260310002820','CFS',None,tmp_path)==d


def test_batch_one_call_and_evidence_validation():
    items=[Item('1','이자',Decimal(1),Decimal(1),'A1',note='현금성자산 이자'),Item('2','비용',Decimal(-1),Decimal(-1),'A2',note='차입금 이자')]
    rows=[Row(item_id=i.id,account_name=i.name,original_amount=str(i.current),category=c,reason='원천 확인',evidence=[{'source_id':'note','quote':q}],paragraphs=[p],confidence=.9,needs_review=False,required_information=[]) for i,c,q,p in zip(items,['투자','재무'],['현금성자산','차입금'],['53','60'])]
    client=Mock();client.responses.parse.return_value=SimpleNamespace(output_parsed=Batch(results=rows))
    result=classify_batch(items,Company('기업',2025),None,'model',client)
    assert client.responses.parse.call_count==1 and result['1'][0]['category']=='투자'
    rows[0].evidence[0].quote='만들어낸 원천'
    result=classify_batch(items,Company('기업',2025),None,'model',client)
    assert result['1'][0] is None and result['2'][0]['category']=='재무'


def test_no_approval_gate(tmp_path):
    source=make(tmp_path)
    from ifrs18.parser import parse,detect
    items,_=parse(source,detect(source)[0]);fx=next(i for i in items if i.name=='외환차익')
    result={'account_name':fx.name,'original_amount':str(fx.current),'category':'영업','reason':'주석 원천 확인','evidence':[{'source_id':'note','quote':'매출채권'}],'paragraphs':['B65'],'confidence':.9,'needs_review':False,'required_information':[]}
    with patch('convert.classify_batch',return_value={fx.id:(result,'')}):
        convert(source,tmp_path/'auto.xlsx',Company('기업',2025,activities='제조업'),download_notes=False)
    w=load_workbook(tmp_path/'auto.xlsx');s=w['계정재분류']
    row=next(r for r in range(2,s.max_row+1) if s.cell(r,2).value=='외환차익')
    assert s.cell(row,14).value=='자동 적용'
    assert '자동 적용' in s.cell(row,20).value
