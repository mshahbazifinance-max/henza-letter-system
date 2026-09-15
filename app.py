from fastapi import FastAPI, Request, Form, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from jinja2 import Environment, FileSystemLoader
from pathlib import Path
import sqlite3, zipfile, re, shutil, os
from datetime import datetime

BASE = Path(__file__).parent
DATA = BASE / 'data'; DATA.mkdir(exist_ok=True)
TEMPLATES = BASE / 'templates'; OUT = DATA / 'letters'; OUT.mkdir(exist_ok=True)
DB = DATA / 'letters.db'
app = FastAPI(title='سامانه مدیریت نامه‌های شرکت')
app.mount('/static', StaticFiles(directory=BASE/'static'), name='static')
env = Environment(loader=FileSystemLoader(BASE/'static'), autoescape=True)

def db():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def init_db():
    c=db(); c.executescript('''
    CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS templates (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, category TEXT NOT NULL, filename TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS letters (id INTEGER PRIMARY KEY AUTOINCREMENT, number TEXT, date TEXT, recipient TEXT, subject TEXT, category TEXT, template_name TEXT, file_path TEXT, created_at TEXT);
    ''')
    defaults={'company_name':'گام آبی فردا (هنزا)','national_id':'10102480408','last_number':'0'}
    for k,v in defaults.items(): c.execute('INSERT OR IGNORE INTO settings VALUES (?,?)',(k,v))
    if not c.execute('SELECT 1 FROM templates').fetchone():
        c.execute('INSERT INTO templates(name,category,filename) VALUES(?,?,?)',('قالب اصلی شرکت','عمومی','قالب-اصلی.docx'))
    c.commit(); c.close()
init_db()

def settings():
    c=db(); rows=c.execute('SELECT key,value FROM settings').fetchall(); c.close(); return {r['key']:r['value'] for r in rows}

def next_number():
    c=db(); n=int(c.execute("SELECT value FROM settings WHERE key='last_number'").fetchone()['value'])+1
    c.execute("UPDATE settings SET value=? WHERE key='last_number'",(str(n),)); c.commit(); c.close(); return str(n)

def replace_all_xml(src, dst, vals):
    with zipfile.ZipFile(src,'r') as zin, zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename.endswith('.xml'):
                text=data.decode('utf-8')
                for k,v in vals.items(): text=text.replace('{{'+k+'}}', str(v or ''))
                data=text.encode('utf-8')
            zout.writestr(item,data)

def build_template_if_needed():
    p=TEMPLATES/'قالب-اصلی.docx'
    # Convert visible metadata values in the supplied company template into tokens.
    tmp=TEMPLATES/'_prepared.docx'
    if tmp.exists(): return tmp
    with zipfile.ZipFile(p,'r') as zin, zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename.endswith('.xml'):
                text=data.decode('utf-8')
                text=text.replace('4110282','{{شماره}}')
                text=text.replace('<w:t>15</w:t></w:r><w:r', '<w:t>{{روز}}</w:t></w:r><w:r',1)
                # Easier robust date replacement across text runs: replace the complete rendered date sequence when present.
                text=text.replace('15</w:t></w:r><w:r', '{{روز}}</w:t></w:r><w:r',1)
                text=text.replace('<w:t>06</w:t></w:r><w:r', '<w:t>{{ماه}}</w:t></w:r><w:r',1)
                text=text.replace('<w:t>1405</w:t></w:r>', '<w:t>{{سال}}</w:t></w:r>',1)
                text=text.replace('پیوست: دارد','پیوست: {{پیوست}}')
                data=text.encode('utf-8')
            zout.writestr(item,data)
    return tmp

def render_letter(number,date,attachment,recipient,recipient_role,subject,body,category,template_name):
    src=build_template_if_needed()
    out=OUT/f'{number} - {subject[:40] or "نامه"}.docx'
    # date is expected yyyy/mm/dd
    parts=date.split('/')
    vals={'شماره':number,'روز':parts[2] if len(parts)>2 else '', 'ماه':parts[1] if len(parts)>1 else '', 'سال':parts[0] if parts else '', 'پیوست':attachment,
          'مخاطب':recipient,'سمت':recipient_role,'موضوع':subject,'متن':body,'نام_شرکت':settings()['company_name'],'شناسه_ملی':settings()['national_id']}
    # Append editable body section to the original template while preserving its design.
    temp=BASE/'data'/'work-template.docx'
    shutil.copy2(src,temp)
    with zipfile.ZipFile(temp,'r') as zin:
        xml=zin.read('word/document.xml').decode('utf-8')
    insert='''<w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:rFonts w:cs="B Nazanin"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:rtl/><w:lang w:bidi="fa-IR"/></w:rPr></w:pPr></w:p>'''
    def p(txt,bold=False):
        b='<w:b/><w:bCs/>' if bold else ''
        return f'<w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:rFonts w:cs="B Nazanin"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:rtl/><w:lang w:bidi="fa-IR"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cs="B Nazanin"/><w:sz w:val="28"/><w:szCs w:val="28"/>{b}</w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>'
    bodyxml=insert+p(f'مخاطب: {recipient}  {recipient_role}')+p(f'موضوع: {subject}',True)+p('با سلام و احترام')
    for line in body.splitlines() or ['']:
        bodyxml += p(line)
    bodyxml += p('با تشکر و احترام')+p(settings()['company_name'])
    xml=xml.replace('</w:body>', bodyxml+'</w:body>')
    with zipfile.ZipFile(temp,'w',zipfile.ZIP_DEFLATED) as zout:
        # cannot write while reading same file; rebuild from original prepared template
        pass
    # Recreate cleanly with modified XML.
    with zipfile.ZipFile(src,'r') as zin, zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename=='word/document.xml': data=xml.encode('utf-8')
            elif item.filename.endswith('.xml'):
                text=data.decode('utf-8')
                for k,v in vals.items(): text=text.replace('{{'+k+'}}',str(v or ''))
                data=text.encode('utf-8')
            zout.writestr(item,data)
    return out

@app.get('/', response_class=HTMLResponse)
async def home(request: Request):
    c=db(); ts=c.execute('SELECT * FROM templates ORDER BY category,name').fetchall(); ls=c.execute('SELECT * FROM letters ORDER BY id DESC LIMIT 10').fetchall(); c.close()
    return env.get_template('index.html').render(settings=settings(), templates=ts, letters=ls)

@app.post('/letters')
async def create_letter(number:str=Form(''), date:str=Form(''), attachment:str=Form('ندارد'), recipient:str=Form(''), recipient_role:str=Form(''), subject:str=Form(''), body:str=Form(''), category:str=Form('عمومی'), template_id:int=Form(1)):
    if not number: number=next_number()
    if not date:
        import jdatetime; date=jdatetime.date.today().strftime('%Y/%m/%d')
    c=db(); t=c.execute('SELECT * FROM templates WHERE id=?',(template_id,)).fetchone(); c.close()
    path=render_letter(number,date,attachment,recipient,recipient_role,subject,body,category,t['name'] if t else 'قالب اصلی')
    c=db(); c.execute('INSERT INTO letters(number,date,recipient,subject,category,template_name,file_path,created_at) VALUES(?,?,?,?,?,?,?,?)',(number,date,recipient,subject,category,t['name'] if t else 'قالب اصلی',str(path),datetime.now().isoformat())); c.commit(); c.close()
    return FileResponse(path,filename=path.name,media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.post('/settings')
async def save_settings(company_name:str=Form(...), national_id:str=Form(...), last_number:str=Form(...)):
    c=db()
    for k,v in [('company_name',company_name),('national_id',national_id),('last_number',last_number)]: c.execute('UPDATE settings SET value=? WHERE key=?',(v,k))
    c.commit(); c.close(); return RedirectResponse('/',status_code=303)

@app.post('/templates')
async def upload_template(name:str=Form(...),category:str=Form(...),file:UploadFile=File(...)):
    safe=re.sub(r'[^\w\-\.آ-ی ]','_',file.filename)
    target=TEMPLATES/safe
    target.write_bytes(await file.read())
    c=db(); c.execute('INSERT INTO templates(name,category,filename) VALUES(?,?,?)',(name,category,safe)); c.commit(); c.close(); return RedirectResponse('/',status_code=303)
