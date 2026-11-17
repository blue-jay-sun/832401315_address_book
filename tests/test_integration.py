import io,json,tempfile,threading,unittest,urllib.request,urllib.error,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import app
from openpyxl import Workbook,load_workbook
class Integration(unittest.TestCase):
 def test_workflow(self):
  with tempfile.TemporaryDirectory() as directory:
   app.DB=Path(directory)/'test.db'
   server=app.ThreadingHTTPServer(('127.0.0.1',0),app.Handler)
   worker=threading.Thread(target=server.serve_forever,daemon=True);worker.start()
   def call(path,data=None,method=None):
    with urllib.request.urlopen(urllib.request.Request('http://127.0.0.1:'+str(server.server_port)+path,data=data,method=method)) as response:return response.read()
   try:
    item={'name':'测试联系人','phones':'13800138000\n13900139000','emails':'one@example.com\ntwo@example.com','social':'微信：demo','addresses':'福州\n北京','notes':'=1+1','favorite':True}
    call('/api/contacts',json.dumps(item).encode(),'POST')
    rows=json.loads(call('/api/contacts'));self.assertEqual(rows[0]['phones'],item['phones']);self.assertTrue(rows[0]['favorite'])
    item.update(id=rows[0]['id'],name='编辑后');call('/api/contacts',json.dumps(item).encode(),'POST')
    excel=call('/api/export');book=load_workbook(io.BytesIO(excel));self.assertEqual(book.active['F2'].data_type,'s')
    self.assertEqual(json.loads(call('/api/import',excel,'POST'))['count'],1)
    bad=Workbook();bad.active.append(app.HEADERS);bad.active.append(['有效']);bad.active.append(['', '123']);out=io.BytesIO();bad.save(out)
    with self.assertRaises(urllib.error.HTTPError):call('/api/import',out.getvalue(),'POST')
    self.assertEqual(len(json.loads(call('/api/contacts'))),2)
    call('/api/contacts/'+str(item['id']),method='DELETE');self.assertEqual(len(json.loads(call('/api/contacts'))),1)
    self.assertIn(b'<html',call('/'))
   finally:server.shutdown();server.server_close();worker.join()
if __name__=='__main__':unittest.main()
