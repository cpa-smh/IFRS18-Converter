"""Excel 손익계산서 → IFRS 18 검토용 손익계산서. 실행: python convert.py --help"""
import argparse
from copy import copy
from dataclasses import asdict
from datetime import datetime
from io import BytesIO
import json
import os
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.properties import CalcProperties
from dotenv import load_dotenv
from ifrs18.models import Company,CATEGORIES
from ifrs18.parser import detect,parse
from ifrs18.engine import classify,RULES
from ifrs18.ai import recommend
from ifrs18.batch_ai import classify_batch
from ifrs18.dart_notes import collect,attach,business_context
import re

ROOT=Path(__file__).resolve().parent


def text(value):
    """회계자료의 텍스트를 Excel 명령으로 실행하지 않는다."""
    value=str(value or '')
    return "'"+value if value.startswith(('=','+','-','@')) else value


def number(value):return float(value) if value is not None else None


def convert(source,output,company,mapping=None,use_ai=True,scale=1,sign_mode='signed',receipt=None,fs_div='CFS',cache_dir=None,download_notes=True,auto_apply=True):
    source=Path(source);output=Path(output)
    if source.resolve()==output.resolve():raise ValueError('원본 보호: 출력 경로를 다르게 지정하세요.')
    if output.exists():raise ValueError('같은 이름의 출력 파일이 이미 있습니다. 새 파일명을 지정하세요.')
    if source.suffix.lower()!='.xlsx':raise ValueError('.xlsx 파일을 사용하세요.')
    if mapping is None:
        found=detect(source)
        if len(found)!=1:raise ValueError('열을 자동 확정할 수 없습니다. --sheet --account-col --current-col --header-row를 지정하세요.')
        mapping=found[0]
    items,ignored=parse(source,mapping,sign_mode,scale)
    prior=bool(mapping.get('previous'))
    # 원본 통합문서를 열어 새 시트만 추가하므로 기존 값·수식·병합·서식을 유지한다.
    w=load_workbook(source)
    business=''
    if download_notes:
        receipts={m for sheet in w for row in sheet for cell in row if isinstance(cell.value,str) for m in re.findall(r'rcpNo=(\d{14})',cell.value)}
        if receipt and receipts and receipts!={receipt}:raise ValueError('입력 재무제표와 요청한 주석 공시의 접수번호가 다릅니다.')
        if not receipt:
            if len(receipts)==1:receipt=receipts.pop()
        if receipt:
            try:
                corpus=collect(receipt,fs_div,os.getenv('OPENDART_API_KEY'),cache_dir or ROOT/'cache')
                audit=''.join(d['text'][:12000] for d in corpus['documents'] if d['role']=='audit')
                if company.name.replace(' ','') not in re.sub(r'\s+','',audit):raise ValueError('주석 감사보고서의 회사명과 입력 기업명이 일치하지 않습니다.')
                attach(items,corpus)
                business=business_context(corpus)
                ignored.append('사용 공시: '+corpus['source_url']+' / '+fs_div+' / 관련 발췌는 계정별 원천자료 칸에 기록')
            except ValueError as exc:ignored.append(str(exc))
        else:ignored.append('주석 자동 수집: 입력 파일에 공시 링크가 없습니다. --rcept-no로 사업보고서 접수번호를 지정하세요.')
    decisions={i.id:classify(i,company) for i in items}
    uncertain=[i for i in items if i.role=='detail' and decisions[i.id].category=='판단 보류']
    batch=classify_batch(uncertain,company,os.getenv('OPENAI_API_KEY'),os.getenv('OPENAI_MODEL',''),business_context=business) if use_ai else {}
    names=('계정재분류','공시IS_IFRS18','검증결과','분류근거')
    if any(n in w.sheetnames for n in names):raise ValueError('이미 변환 시트가 있습니다. 원본 파일을 선택하세요.')
    review,statement,checks,rules=(w.create_sheet(n) for n in names)
    review.append(['원본행 ID','계정과목','당기(원)','전기(원)','기존 범주','규칙 분류','AI 추천','판단 근거','기준서 문단','사용한 원천자료','AI 신뢰도','필요한 증빙·추가자료','사용자 최종분류','검토 확인','반영 범주','출력 구분','원본 위치','행유형','입력 오류','검토 상태'])
    detailrows=[];ocirows=[];netrows=[];metadata=[];audit=[]
    for item in items:
        d=decisions[item.id]; ai,error=batch.get(item.id,(None,''))
        missing=ai['required_information'] if ai else (['발생 원천별 계정명세·관련 주석. 외환손익은 채권/차입금 등 원천과 금액 구분'] if d.category=='판단 보류' else [])
        reason=ai['reason'] if ai else d.reason
        if error:reason+=' / '+error
        evidence=json.dumps(ai['evidence'],ensure_ascii=False) if ai else item.note or item.description or company.activities
        review.append([item.id,text(item.name),number(item.current),number(item.previous),text(item.old_category),d.category,
                       ai['category'] if ai else '',text(reason),', '.join(ai['paragraphs'] if ai else d.refs),text(evidence),ai['confidence'] if ai else None,
                       text('; '.join(missing)),'','미검토',None,d.group,text(item.source),item.role,text('; '.join(item.issues)),None])
        r=review.max_row
        if auto_apply and ai and ai['category'] in CATEGORIES:
            review.cell(r,14,'자동 적용')
        if item.role=='detail':
            detailrows.append(r)
            review.cell(r,15,f'=IF(M{r}<>"",M{r},IF(F{r}<>"판단 보류",F{r},IF(G{r}<>"",G{r},"판단 보류")))')
            review.cell(r,20,f'=IF(OR(S{r}<>"",NOT(ISNUMBER(C{r})){",NOT(ISNUMBER(D"+str(r)+"))" if prior else ""}),"입력 확인",IF(COUNTIFS(\'분류근거\'!$H$2:$H$6,O{r})<>1,"판단 보류",IF(AND(M{r}="",F{r}<>"판단 보류"),"규칙 적용",IF(OR(N{r}="확인",N{r}="자동 적용"),"확인 완료","미검토"))))')
        elif item.role=='oci':
            ocirows.append(r);review.cell(r,15,'OCI');review.cell(r,20,'별도 표시' if item.current is not None and (not prior or item.previous is not None) and not item.issues else '입력 확인')
        else:
            review.cell(r,15,'제외');review.cell(r,20,'통제행')
            if item.role=='net_control':
                netrows.append(r)
                if item.current is None or (prior and item.previous is None) or item.issues:review.cell(r,20,'입력 확인')
        metadata.append({'item':asdict(item),'decision':asdict(d),'ai':ai,'ai_error':error})
    end=review.max_row
    dv=DataValidation(type='list',formula1='"'+','.join((*CATEGORIES,'판단 보류'))+'"');dv.error='목록의 범주를 선택하세요.';dv.showErrorMessage=True;dv.errorStyle='stop';review.add_data_validation(dv);dv.add(f'M2:M{end}')
    approval=DataValidation(type='list',formula1='"미검토,확인,자동 적용"');review.add_data_validation(approval);approval.add(f'N2:N{end}')
    rules.append(['유형','조건/계정','분류','판단 요약','문단','기준 확인일','공식 출처','선택 범주'])
    for idx,c in enumerate(CATEGORIES,2):rules.cell(idx,8,c)
    sourceurl='https://kasb.or.kr/front/board/comm010View.do?seq=2101'
    for rule in RULES['rules']:
        rules.append(['일반 규칙',', '.join(rule['names']),rule['category'],rule['reason'],', '.join(rule['refs']),'2026-11-04',sourceurl])
    for name,rule in RULES['source_rules'].items():
        rules.append(['원천 확인 후 적용',name,rule['category'],rule['reason'],', '.join(rule['refs']),'2026-11-04',sourceurl])
    rules.append(['국내 추가 공시','종전 영업손익과 변경 내역','참고','재분류조정표 및 종전 영업손익은 검토표의 기존 범주를 바탕으로 확인','한132.4~한132.5','2026-11-04',sourceurl])
    statement.append([company.name+' 손익계산서 (IFRS 18 검토용)'])
    statement.append(['기업 업종 / 주된 사업',text(company.industry+' / '+company.activities)])
    statement.append(['상태','초안 — 검증결과 확인 필요'])
    statement.append(['범주 / 계정',company.year,company.year-1 if prior else '전기 미입력'])
    labels={5:'영업 범주',6:'매출액',7:'매출원가',8:'매출총손익',9:'판매비와관리비',10:'기타 영업수익·비용',11:'영업손익',12:'투자 범주',13:'지분법손익',14:'기타 투자수익·비용',15:'재무손익 및 법인세비용 차감전 손익',16:'재무 범주',17:'차입금 관련 손익',18:'리스부채 이자비용',19:'기타 재무수익·비용',20:'법인세비용차감전손익',21:'법인세손익',22:'계속영업손익',23:'중단영업손익',24:'당기순손익',25:'미분류 손익(합계 대사에서 별도 확인)',26:'기타포괄손익',27:'총포괄손익'}
    if company.finances_customers and company.unrelated_policy=='영업':labels[15]='영업·투자범주 손익 합계 (문단 73 추가표시)'
    for r,label in labels.items():statement.cell(r,1,label)
    for col,field in ((2,'C'),(3,'D')):
        if col==3 and not prior:continue
        vals=f"'계정재분류'!${field}$2:${field}${end}";cats=f"'계정재분류'!$O$2:$O${end}";groups=f"'계정재분류'!$P$2:$P${end}";roles=f"'계정재분류'!$R$2:$R${end}"
        def total(category,group=None):
            extra=f', {groups},"{group}"' if group else ''
            return f'SUMIFS({vals},{cats},"{category}"{extra})'
        let='B' if col==2 else 'C'
        formulas={6:total('영업','revenue'),7:total('영업','cost'),8:f'{let}6+{let}7',9:total('영업','sga'),10:f'{total("영업")}-{let}6-{let}7-{let}9',11:f'SUM({let}8:{let}10)',13:total('투자','equity_method'),14:f'{total("투자")}-{let}13',15:f'{let}11+SUM({let}13:{let}14)',17:total('재무','loan_interest'),18:total('재무','lease_interest'),19:f'{total("재무")}-{let}17-{let}18',20:f'{let}15+SUM({let}17:{let}19)',21:total('법인세'),22:f'{let}20+{let}21',23:total('중단영업'),24:f'{let}22+{let}23',25:f'SUMIFS({vals},{roles},"detail")-SUM({let}11,{let}13:{let}14,{let}17:{let}19,{let}21,{let}23)',26:total('OCI'),27:f'{let}24+{let}26'}
        for r,f in formulas.items():statement.cell(r,col,'='+f)
    checks.append(['검증 항목','당기','전기','설명'])
    pending=f'COUNTIFS(\'계정재분류\'!$T$2:$T${end},"미검토")+COUNTIFS(\'계정재분류\'!$T$2:$T${end},"판단 보류")+COUNTIFS(\'계정재분류\'!$T$2:$T${end},"입력 확인")'
    checks.append(['미완료 계정 수','='+pending,None,'근거 검증을 통과한 AI 분류는 자동 적용. 근거 부족·통합 계정은 보류'])
    for col,field in ((2,'C'),(3,'D')):
        if col==3 and not prior:continue
        checks.cell(3,1,'원본 상세 합계 - 원본 순손익')
        checks.cell(4,1,'변환 순손익 - 원본 순손익')
        checks.cell(5,1,'미분류 손익')
        checks.cell(6,1,'완료 여부')
        vals=f"'계정재분류'!${field}$2:${field}${end}";roles=f"'계정재분류'!$R$2:$R${end}";letter='B' if col==2 else 'C'
        control=f'SUMIFS({vals},{roles},"net_control")'
        one=f'COUNTIFS({roles},"net_control")=1'
        checks.cell(3,col,f'=IF({one},SUMIFS({vals},{roles},"detail")-{control},"대사 기준 없음/중복")')
        checks.cell(4,col,f'=IF({one},\'공시IS_IFRS18\'!{letter}24-{control},"대사 기준 없음/중복")')
        checks.cell(5,col,f"='공시IS_IFRS18'!{letter}25")
        checks.cell(6,col,f'=IF(AND(B2=0,ISNUMBER({letter}3),ISNUMBER({letter}4),ABS({letter}3)<0.01,ABS({letter}4)<0.01,ABS({letter}5)<0.01),"검토 완료","초안 / 확인 필요")')
    checks['D3']='중간합계·당기순손익 통제행은 상세 합계에서 제외'
    checks['D4']='OCI는 당기순손익 대사에 포함하지 않음'
    for message in ignored:checks.append(['읽기 안내',None,None,text(message)])
    # 출력 상태는 입력·분류·손익 대사를 직접 보며 검증 시트 결과를 계산 입력으로 사용하지 않는다.
    validnet=f'COUNTIFS(\'계정재분류\'!$R$2:$R${end},"net_control")=1'
    netcur=f'SUMIFS(\'계정재분류\'!$C$2:$C${end},\'계정재분류\'!$R$2:$R${end},"net_control")'
    netprev=f'SUMIFS(\'계정재분류\'!$D$2:$D${end},\'계정재분류\'!$R$2:$R${end},"net_control")'
    statement['B3']=f'=IF(AND({pending}=0,{validnet},ABS(B24-{netcur})<0.01,ABS(B25)<0.01{",ABS(C24-"+netprev+")<0.01,ABS(C25)<0.01" if prior else ""}),"검토 완료","초안 / 확인 필요")'
    # 사용자가 제공한 양식을 참고한 표 구조·서식. 원본 템플릿의 수식은 복사하지 않는다.
    for s in (review,statement,checks,rules):
        s.sheet_view.showGridLines=False;s.freeze_panes='C2' if s!=statement else None
        for row in s:
            for c in row:
                c.font=Font(name='맑은 고딕',size=10);c.alignment=Alignment(vertical='center')
                if isinstance(c.value,(int,float)) or c.data_type=='f':c.number_format='#,##0.00;(#,##0.00);"-"'
        for c in s[4 if s==statement else 1]:
            c.fill=PatternFill('solid',fgColor='173F5F');c.font=Font(name='맑은 고딕',bold=True,color='FFFFFF',size=10)
        s.row_dimensions[4 if s==statement else 1].height=28
        for i in range(1,s.max_column+1):
            from openpyxl.utils import get_column_letter
            s.column_dimensions[get_column_letter(i)].width=23
    for c in ('B','H','J','L'):review.column_dimensions[c].width=48
    review.column_dimensions['I'].width=24;review.column_dimensions['Q'].width=30
    for col in ('M','N'):
        for row in range(2,end+1):review[f'{col}{row}'].fill=PatternFill('solid',fgColor='FFF2CC')
    statement.column_dimensions['A'].width=58;statement.column_dimensions['B'].width=29;statement.column_dimensions['C'].width=25
    statement['A1'].font=Font(name='맑은 고딕',size=14,bold=True);statement['B2'].alignment=Alignment(wrap_text=True)
    statement['B4'].number_format=statement['C4'].number_format='0'
    statement['A4']='범주 / 계정 (단위: 원)'
    checks['B2'].number_format='0'
    statement.row_dimensions[2].height=35
    for r in (8,11,15,20,22,24,27):
        for c in statement[r][:3]:c.font=Font(name='맑은 고딕',size=10,bold=True);c.border=Border(top=Side(style='thin',color='B8C4CF'))
    checks.column_dimensions['A'].width=40;checks.column_dimensions['D'].width=82
    rules.column_dimensions['B'].width=46;rules.column_dimensions['D'].width=70;rules.column_dimensions['G'].width=62
    w.calculation=CalcProperties(fullCalcOnLoad=True,forceFullCalc=True)
    if auto_apply:
        for sheet in (statement,checks):
            for row in sheet:
                for cell in row:
                    if isinstance(cell.value,str):cell.value=cell.value.replace('검토 완료','자동 분류 완료')
    output.parent.mkdir(parents=True,exist_ok=True);w.save(output)
    return {'file':str(output),'accounts':len(detailrows),'ai_calls':sum(x['ai'] is not None or bool(x['ai_error']) for x in metadata),'items':metadata}


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('input',type=Path);p.add_argument('--output',type=Path,default=Path('converted_ifrs18.xlsx'))
    p.add_argument('--company',default='분석기업');p.add_argument('--year',type=int,required=True)
    p.add_argument('--industry',default='제조업');p.add_argument('--activities',required=True)
    p.add_argument('--sheet');p.add_argument('--header-row',type=int);p.add_argument('--account-col');p.add_argument('--current-col');p.add_argument('--previous-col')
    p.add_argument('--no-ai',action='store_true');p.add_argument('--scale',default='1');p.add_argument('--expenses-positive',action='store_true')
    p.add_argument('--rcept-no',help='사업보고서 접수번호 14자리. 입력 공시 링크에서 자동 인식 가능')
    p.add_argument('--fs-div',choices=('CFS','OFS'),default='CFS')
    p.add_argument('--cache-dir',type=Path,default=ROOT/'cache');p.add_argument('--no-documents',action='store_true')
    p.add_argument('--invests-assets',action='store_true');p.add_argument('--invests-financial-assets',action='store_true');p.add_argument('--finances-customers',action='store_true')
    a=p.parse_args(argv);load_dotenv(ROOT/'.env')
    try:
        mapping=None
        if any((a.sheet,a.header_row,a.account_col,a.current_col,a.previous_col)):
            if not all((a.sheet,a.header_row,a.account_col,a.current_col)):raise ValueError('수동 열 지정 시 시트·헤더행·계정열·당기열을 함께 입력하세요.')
            mapping={'sheet':a.sheet,'header_row':a.header_row,'name':a.account_col.upper(),'current':a.current_col.upper()}
            if a.previous_col:mapping['previous']=a.previous_col.upper()
            # 선택한 헤더행에서 주석·설명 열도 인식한다.
            found=[c for c in detect(a.input,a.sheet) if c['header_row']==a.header_row]
            if len(found)==1:mapping={**found[0],**mapping}
        result=convert(a.input,a.output,Company(a.company,a.year,a.industry,a.activities,a.invests_assets,a.invests_financial_assets,a.finances_customers),mapping,not a.no_ai,a.scale,'expenses_positive' if a.expenses_positive else 'signed',a.rcept_no,a.fs_div,a.cache_dir,not a.no_documents)
        print(f'저장: {result["file"]} / 상세 계정 {result["accounts"]}개')
        return 0
    except (ValueError,KeyError) as e:
        print('입력 확인: '+str(e));return 1
    except OSError:
        print('파일 경로·권한을 확인하고 출력 Excel을 닫아 주세요.');return 1


if __name__=='__main__':raise SystemExit(main())
