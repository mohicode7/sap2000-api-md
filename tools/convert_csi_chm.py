import sys, re, json, hashlib, shutil, posixpath
from pathlib import Path
from urllib.parse import unquote, urlsplit, quote
sys.path.insert(0,str(Path(__file__).parent/'chm-tools/lib'))
from bs4 import BeautifulSoup
from markdownify import markdownify
ROOT=Path(__file__).resolve().parents[1]
def run(product,version,chm):
 src=ROOT/'API/_extracted'/f'{product}-{version}'
 out=ROOT/'research/converted'/f'{product}-api-md'; out.mkdir(parents=True,exist_ok=True)
 files=sorted(p for p in src.rglob('*') if p.suffix.lower() in ('.htm','.html'))
 mapping={p.relative_to(src).as_posix().lower():p.relative_to(src).with_suffix('.md').as_posix() for p in files}
 titles={}; inventory=[]; broken=[]
 def soup(p):
  raw=p.read_bytes(); enc='cp1252' if b'windows-1252' in raw[:2000].lower() else 'utf-8-sig'
  return BeautifulSoup(raw.decode(enc,errors='strict'),'html.parser')
 for p in files:
  doc=soup(p); rel=p.relative_to(src).as_posix(); target=out/mapping[rel.lower()]; target.parent.mkdir(parents=True,exist_ok=True)
  title=doc.title.get_text(' ',strip=True) if doc.title else p.stem; titles[rel.lower()]=title
  body=doc.select_one('#TopicContent') or doc.body or doc
  # Sandcastle inserts punctuation and array notation dynamically. Materialize the neutral language form.
  for script in body.find_all('script'):
   for ident,values in re.findall(r'AddLanguageSpecificTextSet\("([^?]+)\?([^"\n]+)"\)',script.get_text()):
    span=body.find(id=ident); variants=dict(v.split('=',1) for v in values.split('|') if '=' in v)
    if span is not None: span.string=variants.get('cs',variants.get('nu',next(iter(variants.values()),'')))
  for e in body.select('script,style,input,.codeSnippetContainerTabs,.codeSnippetToolBar,.collapseToggle,.pageFooter'): e.decompose()
  for e in body.select('.collapsibleRegionTitle'):
   e.name='h2'; e.attrs={}
  for table in list(body.select('table.titleTable')):
   heading=table.find('h1')
   if heading: table.replace_with(heading.extract())
  for para in list(body.select('p.Coding')):
   for br in para.find_all('br'): br.replace_with('\n')
   para.name='pre'; para.string=para.get_text().replace('\xa0',' '); para.attrs={}
  for ol in body.find_all('ol'):
   first=ol.find('li')
   if first and first.get('value'): ol['start']=first['value']
  code_count=len(body.find_all('pre'))
  anchors=[]
  if product=='etabs': body.insert(0,BeautifulSoup('<a id="PageHeader"></a>','html.parser'))
  for e in body.find_all(True):
   ident=e.get('id') or (e.get('name') if e.name=='a' else None)
   if ident: anchors.append(ident); e.insert_before(BeautifulSoup('<a id="'+str(ident).replace('"','&quot;')+'"></a>','html.parser'))
  for e,attr in [(e,'href') for e in body.find_all('a',href=True)]+[(e,'src') for e in body.find_all('img',src=True)]:
   url=e[attr]; u=urlsplit(url)
   if u.scheme or u.netloc or not u.path: continue
   resolved=posixpath.normpath(posixpath.join(posixpath.dirname(rel),unquote(u.path)))
   dest=mapping.get(resolved.lower())
   if dest is None:
    asset=src/resolved
    if asset.is_file():
     dest='source-assets/'+resolved; (out/dest).parent.mkdir(parents=True,exist_ok=True); shutil.copy2(asset,out/dest)
    else: broken.append({'source':rel,'link':url}); continue
   e[attr]=quote(posixpath.relpath(dest,posixpath.dirname(mapping[rel.lower()]) or '.'),safe='/._-')+('#'+u.fragment if u.fragment else '')
  # Keep explicit anchor elements; Markdown otherwise drops empty named anchors.
  placeholders={}
  for i,e in enumerate(body.select('a[id]')):
   token=f'CSIANCHOR{i}TOKEN'; placeholders[token]=str(e); e.replace_with(token)
  for i,e in enumerate(body.find_all('pre')):
   token=f'CSICODE{i}TOKEN'; code=e.get_text(); fence='`'*max(3,max((len(x) for x in re.findall(r'`+',code)),default=0)+1); placeholders[token]='\n\n'+fence+'\n'+code.strip('\r\n')+'\n'+fence+'\n\n'; e.replace_with(token)
  text=markdownify(str(body),heading_style='ATX',bullets='-',code_language='',strip=['div','span'],wrap=False)
  for token,value in placeholders.items(): text=text.replace(token,value)
  if product=='sap2000': text=re.sub(r'\n```\s*\n```\n','\n',text)
  target.write_text(text.replace('\r\n','\n').strip()+'\n',encoding='utf-8')
  inventory.append({'source':rel,'markdown':mapping[rel.lower()],'title':title,'code_blocks':code_count,'anchors':anchors})
 toc=next(src.glob('*.hhc')); nav=['# Table of contents','']
 for obj in soup(toc).find_all('object'):
  params={x.get('name','').lower():x.get('value','') for x in obj.find_all('param')}; name=params.get('name'); local=params.get('local','').replace('\\','/')
  if not name: continue
  depth=max(0,len(obj.find_parents('ul'))-1); dest=mapping.get(unquote(local).lower())
  nav.append('  '*depth+'- '+(f'[{name}]({quote(dest,safe="/._-")})' if dest else name))
 (out/'TOC.md').write_text('\n'.join(nav)+'\n',encoding='utf-8')
 chmp=ROOT/chm
 provenance={'product':product,'installed_product_version':version,'source_chm':chmp.name,'sha256':hashlib.sha256(chmp.read_bytes()).hexdigest(),'html_topics':len(files),'converted_topics':len(inventory),'unresolved_source_links':broken}
 (out/'manifest.json').write_text(json.dumps({'provenance':provenance,'topics':inventory},indent=2),encoding='utf-8')
 (out/'README.md').write_text(f'# {product.upper()} API documentation\n\nMarkdown conversion of the CSI help shipped with {product.upper()} {version}.\n\n[Browse the original table of contents](TOC.md). Each HTML topic has its own Markdown page, retaining the original source paths.\n\n## Provenance\n\nSource: `{chmp.name}`. SHA-256: `{provenance["sha256"]}`.\n\n{len(files)} HTML topics converted. See [manifest](manifest.json) for source mapping and validation details. Product version and API assembly version are distinct. Documentation content and examples belong to Computers and Structures, Inc.; conversion does not grant additional rights. Historical examples are retained as written and can describe obsolete features; consult release notes and breaking changes.\n',encoding='utf-8')
 print(json.dumps(provenance))
if len(sys.argv)==1 or sys.argv[1]=='etabs': run('etabs','23.3.2','API/ETABS/v23.3.2.4670/CSI API ETABS v1.chm')
if len(sys.argv)==1 or sys.argv[1]=='sap2000': run('sap2000','27.1.0','API/SAP2000/v27.1.0.3560/CSI_OAPI_Documentation.chm')
