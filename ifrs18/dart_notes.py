"""공시 원문 ZIP 1회 수집 → 연결/별도 주석 선택 → 계정별 관련 부분 검색."""
from html.parser import HTMLParser
from io import BytesIO
from pathlib import Path
from urllib.request import urlopen
from urllib.parse import urlencode
from zipfile import ZipFile,BadZipFile
import json,re,hashlib


class Reader(HTMLParser):
    def __init__(self):super().__init__(convert_charrefs=True);self.parts=[]
    def handle_data(self,data):self.parts.append(data)
    def handle_endtag(self,tag):
        if tag.lower() in ('p','tr','table','title','section-1','section-2','section-3'):self.parts.append('\n')
        elif tag.lower() in ('td','th','tu'):self.parts.append(' | ')


def parse_zip(payload,receipt,basis):
    if not payload.startswith(b'PK'):raise ValueError('공시 원문 조회 실패: 인증키·접수번호·API 한도를 확인하세요.')
    try:
        with ZipFile(BytesIO(payload)) as z:
            suffix='_00761.xml' if basis=='CFS' else '_00760.xml'
            audit=[n for n in z.namelist() if n.lower().endswith(suffix)]
            if len(audit)!=1:raise ValueError('같은 공시의 연결/별도 감사보고서 원문을 특정할 수 없습니다.')
            names=[n for n in z.namelist() if n.lower().endswith('.xml') and not re.search(r'_\d{5}\.xml$',n.lower())]+audit
            if sum(z.getinfo(n).file_size for n in names)>80_000_000:raise ValueError('공시 원문 크기 제한 초과')
            documents=[]
            for name in names:
                data=z.read(name)
                try:decoded=data.decode('utf-8-sig')
                except UnicodeDecodeError:decoded=data.decode('cp949')
                reader=Reader();reader.feed(decoded)
                content=re.sub(r'[ \t\r]+',' ',''.join(reader.parts))
                content=re.sub(r'\n\s*\n','\n',content)
                documents.append({'file':name,'role':'audit' if name in audit else 'business_report','text':content})
        return {'schema_version':1,'receipt':receipt,'basis':basis,'source_url':'https://dart.fss.or.kr/dsaf001/main.do?rcpNo='+receipt,'documents':documents}
    except (BadZipFile,KeyError,UnicodeError):raise ValueError('공시 ZIP/문자 인코딩 오류') from None


def collect(receipt,basis,key,cache_dir):
    if not re.fullmatch(r'\d{14}',receipt) or basis not in ('CFS','OFS'):raise ValueError('접수번호/연결·별도 구분 오류')
    cache=Path(cache_dir)/f'{receipt}_{basis}.json'
    if cache.exists():
        d=json.loads(cache.read_text(encoding='utf-8'))
        if d.get('schema_version')==1 and d.get('receipt')==receipt and d.get('basis')==basis:return d
        raise ValueError('공시 캐시 식별정보 불일치')
    if not re.fullmatch(r'[A-Za-z0-9]{40}',key or ''):raise ValueError('OPENDART_API_KEY에 발급받은 40자리 키가 필요합니다.')
    url='https://opendart.fss.or.kr/api/document.xml?'+urlencode({'crtfc_key':key,'rcept_no':receipt})
    try:
        with urlopen(url,timeout=30) as response:payload=response.read(80_000_001)
        if len(payload)>80_000_000:raise ValueError('공시 다운로드 크기 제한 초과')
    except Exception:raise ValueError('공시 다운로드 실패. 인증키·접수번호·네트워크를 확인하세요.') from None
    d=parse_zip(payload,receipt,basis);cache.parent.mkdir(parents=True,exist_ok=True)
    cache.write_text(json.dumps(d,ensure_ascii=False),encoding='utf-8');return d


def sections(corpus):
    """번호가 붙은 주석의 시작부터 다음 주석 직전까지 표·설명을 함께 보존한다."""
    result=[]
    for doc in corpus['documents']:
        body=doc['text']
        if doc['role']=='audit':
            headings=list(re.finditer(r'(?m)^\s*(\d{1,2})[.]\s*([^\n|]{2,90})$',body))
            # 문장·표 셀·괄호 소항목은 주석 제목으로 취급하지 않는다.
            headings=[m for m in headings if not m.group(2).strip().endswith(('다.','니다.'))]
            sequence=[];last=0;limit=len(body)
            for m in headings:
                number=int(m.group(1))
                if last and number==1:
                    limit=m.start();break  # 주석 뒤 외부감사 실시내용의 번호가 다시 시작됨
                if number==last+1:
                    sequence.append(m);last=number
            headings=sequence
            for n,m in enumerate(headings):
                end=headings[n+1].start() if n+1<len(headings) else limit
                result.append({'id':doc['file']+':note:'+m.group(1)+':'+str(m.start()),
                    'number':m.group(1),'title':m.group(2).strip(),
                    'location':('연결' if corpus['basis']=='CFS' else '별도')+' 주석 '+m.group(1)+' '+m.group(2).strip(),
                    'text':body[m.start():end], 'file':doc['file'],'role':'audit','start':m.start()})
            if not headings:
                result.append({'id':doc['file']+':whole','number':'','title':'감사보고서 전체 (주석 경계 인식 실패)',
                    'location':doc['file']+' 전체','text':body,'file':doc['file'],'role':'audit','start':0})
        else:
            chapters=list(re.finditer(r'(?m)^\s*(?:II|Ⅱ)\s*[.]?\s*사업의\s*내용[^\n]*',body))
            if not chapters:continue
            start=chapters[-1].start()
            stop=re.search(r'(?m)^\s*(?:III|Ⅲ)\s*[.]?\s*재무에\s*관한',body[start:])
            chapter=body[start:start+stop.start()] if stop else body[start:]
            marks=list(re.finditer(r'(?m)^\s*\d{1,2}[.]\s*[^\n|]{2,90}$',chapter))
            bounds=[0]+[m.start() for m in marks if m.start()>0]+[len(chapter)]
            for a,b in zip(bounds,bounds[1:]):
                part=chapter[a:b]
                title=part.strip().split('\n')[0]
                result.append({'id':doc['file']+':business:'+str(start+a),'number':'','title':title,
                    'location':'사업보고서 II. 사업의 내용 / '+title,'text':part,'file':doc['file'],'role':'business_report','start':start+a})
    return result


def references(content):
    refs=set()
    for match in re.finditer(r'주석\s*([0-9]{1,2}(?:\s*(?:,|ㆍ|·|및|와|과)\s*[0-9]{1,2})*)',content):
        refs.update(re.findall(r'\d+',match.group(1)))
    return refs


def render_section(section,corpus):
    return '[SOURCE '+section['id']+' | '+section['location']+' | '+corpus['source_url']+']\n'+section['text']+'\n[/SOURCE]'


def related(item,corpus):
    indexed=sections(corpus); selected={}
    terms={item.name}
    if '금융' in item.name:terms.update(('금융수익','금융비용','금융손익','금융상품'))
    if any(t in item.name for t in ('기타','영업외')):terms.update(('기타수익','기타비용','기타손익'))
    if any(t in item.name for t in ('외환','환산')):terms.update(('외환','환율','위험관리','금융상품'))
    explicit=references(item.note)
    if re.fullmatch(r'[\d,\s]+',item.note.strip()):explicit.update(re.findall(r'\d+',item.note))
    for section in indexed:
        title=re.sub(r'\s+','',section['title'])
        if (section['role']=='audit' and (section['id'].endswith(':whole') or section['number'] in explicit or any(re.sub(r'\s+','',t) in title for t in terms)
                or any(t in title for t in ('회계정책','일반사항','재무제표작성','작성기준')))):
            selected[section['id']]=section
    # 본표에서 정확히 일치하는 계정의 주석 번호도 따라간다.
    for doc in corpus['documents']:
        if doc['role']!='audit':continue
        for line in doc['text'].splitlines():
            cells=[c.strip() for c in line.split('|')]
            if cells and re.sub(r'\s+','',cells[0])==re.sub(r'\s+','',item.name):
                for cell in cells[1:]:
                    if re.fullmatch(r'\d{1,2}(?:\s*,\s*\d{1,2})*',cell):explicit.update(re.findall(r'\d+',cell))
    while True:
        needed=explicit | set().union(*(references(s['text']) for s in selected.values()))
        extra=[s for s in indexed if s['role']=='audit' and s['number'] in needed and s['id'] not in selected]
        if not extra:break
        selected.update((s['id'],s) for s in extra)
    return list(selected.values())


def attach(items,corpus):
    for item in items:
        if item.role!='detail':continue
        matched=related(item,corpus)
        if matched:item.note=(item.note+'\n'+'\n\n'.join(render_section(s,corpus) for s in matched)).strip()
    return items


def business_context(corpus):
    return '\n\n'.join(render_section(s,corpus) for s in sections(corpus) if s['role']=='business_report')


def locations(item,evidence,business=''):
    """AI가 만든 번호 대신 실제 인용문이 포함된 원문의 위치를 표시한다."""
    result=[]
    for e in evidence:
        source={'note':item.note,'description':item.description,'activities':business}.get(e['source_id'],'')
        matches=list(re.finditer(r'\[SOURCE ([^\n]+)\]\n(.*?)\n\[/SOURCE\]',source,re.S))
        hits=[' | '.join(m.group(1).split(' | ')[1:]) for m in matches if e['quote'] in m.group(2)]
        if matches and not hits:raise ValueError('引用が原文の本文にありません')
        result.extend(hits or ['入力資料 / '+e['source_id']])
    return list(dict.fromkeys(result))
