"""원본 수식은 건드리지 않는다. 캐시 없는 금액 수식은 사용자 확인 대상으로 남긴다."""
from pathlib import Path
from io import BytesIO
import hashlib
import re
import pandas as pd
from openpyxl import load_workbook
from .models import Item,money
from .engine import normalize

HEADER_NAMES={'name':['계정과목','계정명','계정','account','항목'], 'current':['당기','당기금액','금액','current'],
              'previous':['전기','전기금액','previous'], 'old_category':['기존범주','기존구분','기존영업구분'],
              'description':['설명','계정설명','세부명세'], 'note':['주석','관련주석','근거'],
              'source_type':['원천유형','source_type'], 'role':['행유형','role']}
SUBTOTALS={'매출총이익','매출총손익','영업이익','영업손익','영업손실','법인세비용차감전손익','법인세비용차감전순손익','계속영업손익','총포괄손익','기타포괄손익합계'}
NET={'당기순이익','당기순손익','당기순손실','당기순이익(손실)'}


def workbook_info(path):
    w=load_workbook(path,data_only=False)
    return [{'sheet':s.title,'rows':s.max_row,'columns':s.max_column,'merges':[str(x) for x in s.merged_cells.ranges],
             'formulas':sum(c.data_type=='f' for row in s for c in row)} for s in w]


def detect(path,sheet=None):
    w=load_workbook(path,data_only=False)
    candidates=[]
    for s in ([w[sheet]] if sheet else w):
        for row in s.iter_rows(max_row=min(s.max_row,40)):
            mapping={}
            for key,aliases in HEADER_NAMES.items():
                hits=[c.column_letter for c in row if c.value is not None and normalize(str(c.value)).lower() in [normalize(a) for a in aliases]]
                if len(hits)==1:mapping[key]=hits[0]
            if 'name' in mapping and 'current' in mapping:
                candidates.append({'sheet':s.title,'header_row':row[0].row,**mapping})
    return candidates


def parse(path,mapping,sign_mode='signed',scale=1):
    w=load_workbook(path,data_only=False);cached=load_workbook(path,data_only=True)
    s=w[mapping['sheet']];cs=cached[s.title];items=[];excluded=[]
    scale=money(scale)
    if scale is None or scale<=0: raise ValueError('단위 배수는 0보다 큰 숫자로 지정하세요.')
    columns=[v for k,v in mapping.items() if k not in ('sheet','header_row') and v]
    if len(columns)!=len(set(columns)): raise ValueError('서로 다른 항목에 같은 열을 지정할 수 없습니다.')
    for r in range(int(mapping['header_row'])+1,s.max_row+1):
        def get(key):
            col=mapping.get(key)
            return s[f'{col}{r}'].value if col else None
        name=get('name')
        if name is None or not str(name).strip(): continue
        name=str(name).strip();clean=normalize(re.sub(r'^[IVXⅠⅡⅢⅣⅤⅥⅦ\d.\s]+','',name))
        role=str(get('role') or '').strip()
        if not role:
            role='net_control' if clean in NET else 'subtotal' if clean in SUBTOTALS else 'oci' if '기타포괄' in clean else 'detail'
        if role not in ('detail','oci','subtotal','net_control','ignore'): raise ValueError(f'{r}행의 행유형이 잘못되었습니다.')
        issues=[];values=[]
        for key in ('current','previous'):
            col=mapping.get(key)
            if not col: values.append(None);continue
            cell=s[f'{col}{r}'];value=cs[cell.coordinate].value if cell.data_type=='f' else cell.value
            number=money(value)
            if cell.data_type=='f' and number is None:issues.append(f'{cell.coordinate}: Excel 수식 계산값 없음. Excel에서 재계산·저장 필요')
            elif value is not None and number is None:issues.append(f'{cell.coordinate}: 금액 형식 확인 필요')
            if sign_mode=='expenses_positive' and number is not None and number>0 and any(t in clean for t in ('비용','원가','손실','판매비와관리비')):
                number=-number
            values.append(number*scale if number is not None else None)
        if all(v is None for v in values) and not issues and role=='detail':
            excluded.append(f'{s.title}!{r}: 금액 없는 제목행 {name}');continue
        source=f'{s.title}!{mapping["name"]}{r}'
        uid=hashlib.sha256(f'{source}|{name}'.encode()).hexdigest()[:16]
        items.append(Item(uid,name,*values,source,str(get('old_category') or ''),str(get('description') or ''),str(get('note') or ''),str(get('source_type') or 'unknown'),role,issues))
    if not items:raise ValueError('분석할 계정이 없습니다. 시트·헤더·열을 확인하세요.')
    # 정규화된 레코드를 pandas로 검사. 같은 ID는 원본 행 중복을 뜻한다.
    frame=pd.DataFrame([{'id':i.id,'name':i.name} for i in items])
    if frame['id'].duplicated().any():raise ValueError('중복 원본 행 ID')
    return items,excluded
