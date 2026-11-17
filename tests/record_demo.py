import sys,tempfile,threading,json,io
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import app
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
root=app.ROOT/'delivery'
with tempfile.TemporaryDirectory() as directory:
 app.DB=Path(directory)/'demo.db'
 server=app.ThreadingHTTPServer(('127.0.0.1',0),app.Handler)
 thread=threading.Thread(target=server.serve_forever,daemon=True);thread.start()
 try:
  with sync_playwright() as p:
   browser=p.chromium.launch()
   context=browser.new_context(viewport={'width':1440,'height':950},record_video_dir=str(root/'videos'),record_video_size={'width':1440,'height':950},accept_downloads=True)
   page=context.new_page();errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
   page.goto(f'http://127.0.0.1:{server.server_port}');page.wait_for_timeout(600)
   for name,phone,favorite in [('课程助教（演示）','13800000001\n13800000002',True),('项目伙伴（演示）','13900000001',False),('图书馆（演示）','0591-00000000',False)]:
    page.locator('#add').click()
    for k,v in {'name':name,'phones':phone,'emails':'demo@example.com\nwork@example.com','social':'微信：demo-contact','addresses':'福州大学（演示地址）','notes':'仅用于课程演示，非真实联系方式'}.items():page.locator(f'[name={k}]').fill(v)
    if favorite:page.locator('[name=favorite]').check()
    page.get_by_role('button',name='保存联系人').click();page.wait_for_timeout(400)
   page.screenshot(path=str(root/'screenshots/01-main.png'),full_page=True)
   page.locator('#favorites').click();assert page.locator('.card').count()==1;page.screenshot(path=str(root/'screenshots/02-favorites.png'))
   page.get_by_role('button',name='编辑',exact=True).click();page.screenshot(path=str(root/'screenshots/03-editor.png'));page.locator('[name=notes]').fill('编辑后的备注');page.get_by_role('button',name='保存联系人').click();page.wait_for_timeout(300)
   page.locator('#all').click();page.locator('#search').fill('13800000002');assert page.locator('.card').count()==1;page.wait_for_timeout(700);page.locator('#search').fill('')
   with page.expect_download() as download:page.get_by_text('↓ 导出 Excel',exact=True).click()
   download.value.save_as(str(root/'通讯录示例.xlsx'))
   page.locator('#file').set_input_files(str(root/'通讯录示例.xlsx'));page.wait_for_timeout(600);assert page.locator('.card').count()==6
   page.on('dialog',lambda dialog:dialog.accept());page.locator('.card').first.get_by_role('button',name='删除').click();page.wait_for_timeout(400);assert page.locator('.card').count()==5
   page.reload();page.wait_for_timeout(500);assert page.locator('.card').count()==5
   page.set_viewport_size({'width':390,'height':844});page.screenshot(path=str(root/'screenshots/04-mobile.png'),full_page=True);assert page.evaluate('document.documentElement.scrollWidth <= innerWidth')
   assert not errors,errors
   video=page.video;context.close();video.save_as(str(root/'完整功能演示.webm'));browser.close()
   (root/'界面验收.txt').write_text('PASS：浏览器新增、多值填写、收藏筛选、编辑、搜索、Excel 下载与导入、删除、刷新持久化、390px 移动布局；无 JavaScript 页面错误。使用隔离演示数据库，无真实个人数据。',encoding='utf-8')
 finally:server.shutdown();server.server_close();thread.join()
