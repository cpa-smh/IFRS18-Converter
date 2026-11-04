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


def attach(items,corpus):
    """각 발췌에 문서명·문자 위치·공시 링크를 붙인다. 전체 보고서는 AI에 보내지 않는다."""
    chunks=[]
    for doc in corpus['documents']:
        # 본문의 별도 주석과 연결 주석이 함께 검색되는 것을 막는다.
        # 계정의 회계 증거는 선택한 연결/별도 감사보고서에서만 가져온다.
        if doc['role']!='audit':continue
        for start in range(0,len(doc['text']),2700):
            part=doc['text'][start:start+3000]
            chunks.append((doc['file'],start,part,doc['role']))
    for item in items:
        if item.role!='detail':continue
        terms=[item.name]
        if '금융' in item.name:terms+=['금융수익','금융비용','이자비용','외환차이']
        if '기타' in item.name:terms+=['기타수익','기타비용','처분이익','처분손실']
        if any(t in item.name for t in ('외환','환산')):terms+=['외환차이','매출채권','차입금']
        ranked=sorted(chunks,key=lambda c:(sum(c[2].count(t)*(5 if t==item.name else 1) for t in terms),c[3]=='audit'),reverse=True)
        selected=[c for c in ranked[:2] if any(t in c[2] for t in terms)]
        if selected:
            extracts='\n\n'.join(f'[{f} 문자 {start}:{start+len(part)}]\n{part}' for f,start,part,_ in selected)
            item.note=(item.note+'\n'+corpus['source_url']+'\n'+extracts).strip()
    return items


def business_context(corpus):
    for doc in corpus['documents']:
        if doc['role']!='business_report':continue
        matches=list(re.finditer(r'(?:II|Ⅱ)\s*[.]?\s*사업의\s*내용',doc['text']))
        if matches:
            start=matches[-1].start()
            return f'[{doc["file"]} 문자 {start}:{start+3000}]\n'+corpus['source_url']+'\n'+doc['text'][start:start+3000]
    return ''
