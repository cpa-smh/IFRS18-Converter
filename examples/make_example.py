"""실제 기업과 무관한 입력·기대 결과. 프로젝트 루트에서 실행."""
from pathlib import Path
import sys
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from convert import convert
from ifrs18.models import Company


def make(folder):
    folder=Path(folder);folder.mkdir(parents=True,exist_ok=True)
    w=Workbook();s=w.active;s.title='기존손익'
    s.append(['계정과목','당기','전기','기존범주','설명','주석','원천유형'])
    rows=[['매출액',1000,900,'영업','제품 판매','','unknown'],['매출원가',-600,-550,'영업','판매 제품 원가','','unknown'],
          ['매출총손익',400,350,'','','','unknown'],['판매비와관리비',-200,-180,'영업','판매 및 일반관리비','','unknown'],
          ['영업손익',200,170,'','','','unknown'],['유형자산처분이익',20,10,'영업외','생산설비 매각','다른 생산자원과 함께 사용한 설비의 매각이익','operating_ppe'],
          ['지분법이익',30,20,'영업외','관계기업 지분법 손익','','equity_method'],['차입금이자비용',-40,-30,'영업외','일반 운영자금 차입 이자','고객 금융제공과 관계없는 차입금','loan_funding'],
          ['외환차익',10,8,'영업외','','매출채권에서 발생한 환차익이다. 명세 검토 필요','unknown'],
          ['법인세비용',-5,-4,'법인세','계속영업 법인세','','unknown'],['당기순손익',215,174,'','','','unknown'],
          ['기타포괄손익',7,5,'OCI','세후 OCI','','unknown']]
    for row in rows:s.append(row)
    for c in s[1]:c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='173F5F')
    for col in ('A','E','F'):s.column_dimensions[col].width=46
    p=folder/'input_example.xlsx';w.save(p)
    convert(p,folder/'output_example.xlsx',Company('가상제조',2025,activities='제품 제조 및 판매'),use_ai=False)
    return p


if __name__=='__main__':make(ROOT/'examples')
