import io
from contextlib import closing
import json
import sqlite3
from pathlib import Path
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent
DB = ROOT / 'contacts.db'
HEADERS = ['姓名', '电话', '邮箱', '社交账号', '地址', '备注', '收藏']
KEYS = ['name', 'phones', 'emails', 'social', 'addresses', 'notes', 'favorite']

def connection():
    db = sqlite3.connect(DB)
    db.row_factory = sqlite3.Row
    db.execute('CREATE TABLE IF NOT EXISTS contacts (id INTEGER PRIMARY KEY, data TEXT NOT NULL)')
    return db

def normalize(data):
    item = {key: str(data.get(key, '') or '').strip() for key in KEYS[:-1]}
    if not item['name']:
        raise ValueError('姓名不能为空')
    if any(len(v) > 10000 for v in item.values()):
        raise ValueError('单个字段不能超过 10000 字符')
    item['favorite'] = data.get('favorite') in (True, 1, '1', '是', 'true')
    return item

class Handler(BaseHTTPRequestHandler):
    def send(self, data, status=200, mime='application/json; charset=utf-8', filename=None):
        if not isinstance(data, bytes):
            data = json.dumps(data, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header('Content-Type', mime)
        self.send_header('Content-Length', str(len(data)))
        if filename:
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = urlparse(self.path).path
        if path in ('/', '/style.css', '/app.js'):
            name = 'index.html' if path == '/' else path[1:]
            mime = {'index.html': 'text/html; charset=utf-8', 'style.css': 'text/css', 'app.js': 'text/javascript; charset=utf-8'}[name]
            return self.send((ROOT / 'static' / name).read_bytes(), mime=mime)
        if path not in ('/api/contacts', '/api/export'):
            return self.send({'error': '页面不存在'}, 404)
        with closing(connection()) as db, db:
            items = [dict(json.loads(r['data']), id=r['id']) for r in db.execute('SELECT * FROM contacts ORDER BY id DESC')]
        if path == '/api/contacts':
            return self.send(items)
        wb = Workbook()
        ws = wb.active
        ws.title = '通讯录'
        ws.append(HEADERS)
        for item in items:
            ws.append([('是' if item[k] else '否') if k == 'favorite' else item[k] for k in KEYS])
            for cell in ws[ws.max_row]:
                cell.data_type = 's'
        ws.freeze_panes = 'A2'
        for col in 'ABCDEFG':
            ws.column_dimensions[col].width = 26
        out = io.BytesIO()
        wb.save(out)
        self.send(out.getvalue(), mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename='contacts.xlsx')

    def do_POST(self):
        try:
            size = int(self.headers.get('Content-Length', 0))
            if size > 5 * 1024 * 1024:
                raise ValueError('文件不能超过 5MB')
            raw = self.rfile.read(size)
            path = urlparse(self.path).path
            if path == '/api/import':
                wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
                rows = iter(wb.active.values)
                if list(next(rows, ())) != HEADERS:
                    raise ValueError('表头不匹配，请先导出 Excel 作为模板')
                items = []
                for number, row in enumerate(rows, 2):
                    if number > 10001:
                        raise ValueError('最多导入 10000 条联系人')
                    if not any(v is not None for v in row):
                        continue
                    try:
                        items.append(normalize(dict(zip(KEYS, row))))
                    except ValueError as error:
                        raise ValueError(f'第 {number} 行：{error}')
                with closing(connection()) as db, db:
                    db.executemany('INSERT INTO contacts(data) VALUES (?)', [(json.dumps(x, ensure_ascii=False),) for x in items])
                return self.send({'count': len(items)})
            if path != '/api/contacts':
                return self.send({'error': '接口不存在'}, 404)
            data = json.loads(raw)
            item = normalize(data)
            with closing(connection()) as db, db:
                if data.get('id'):
                    cursor = db.execute('UPDATE contacts SET data=? WHERE id=?', (json.dumps(item, ensure_ascii=False), int(data['id'])))
                    if not cursor.rowcount:
                        return self.send({'error': '联系人不存在'}, 404)
                else:
                    db.execute('INSERT INTO contacts(data) VALUES (?)', (json.dumps(item, ensure_ascii=False),))
            self.send({'ok': True})
        except Exception as error:
            self.send({'error': str(error)}, 400)

    def do_DELETE(self):
        try:
            path = urlparse(self.path).path
            if not path.startswith('/api/contacts/'):
                return self.send({'error': '接口不存在'}, 404)
            with closing(connection()) as db, db:
                db.execute('DELETE FROM contacts WHERE id=?', (int(path.rsplit('/', 1)[1]),))
            self.send({'ok': True})
        except ValueError:
            self.send({'error': '无效编号'}, 400)

if __name__ == '__main__':
    connection().close()
    print('Address book: http://127.0.0.1:8000', flush=True)
    ThreadingHTTPServer(('127.0.0.1', 8000), Handler).serve_forever()
