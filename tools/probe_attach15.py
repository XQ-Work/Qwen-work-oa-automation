# -*- coding: utf-8 -*-
"""探附件15：一行两文件 + 显式 startUpload 推队列。"""
import sys
import time
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

ROOT = paths.DATA / "OA附件库/文档整理"
F1 = str(ROOT / "整理中/文档/比质比价报告单（变频器）.pdf")
F2 = str(ROOT / "整理中/文档/供应商报价单及资质（变频器）.pdf")
OUT = paths.STATE / "probe_attach15.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


STARTQ = r"""
(function(){var n=0;
  if(window.SWFUpload&&SWFUpload.instances){for(var k in SWFUpload.instances){
    try{var u=SWFUpload.instances[k]; var q=u.getStats().files_queued; if(q>0){u.startUpload();n++;}}catch(e){}}}
  return 'startUpload on '+n;})()
"""
INFO = r"""
(function(){var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
  var h=tbl?tbl.querySelector('input[name=field54276_0]'):null;
  return 'docid='+(h?h.value:'NA')+' | '+(tbl?tbl.textContent.replace(/\s+/g,' ').slice(0,120):'');})()
"""


def main():
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            base = auth.login(page)
            page.goto(base + m.ADD, wait_until="domcontentloaded", timeout=40000)
            time.sleep(4)
            tgt = m.find_annex_frame(page)
            tgt.locator("button[id='$addbutton0$']").first.click(force=True, timeout=4000)
            time.sleep(2)
            tgt.evaluate(r"""(function(){var tds=document.querySelectorAll('td'),tbl;
              for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
              tbl.querySelectorAll('input[type=file]').forEach(function(e){e.setAttribute('data-upl','1');});})()""")
            fi = tgt.locator("input[type=file][data-upl='1']")
            fi.last.set_input_files([F1, F2], timeout=8000)
            time.sleep(2)
            w("塞两文件后先 startUpload: " + tgt.evaluate(STARTQ))
            for t in range(12):
                time.sleep(3)
                info = tgt.evaluate(INFO)
                w(f"[{(t+1)*3}s] {info}")
                if re.search(r"docid=\d", info):
                    # 再推一次队列确保第二个也上
                    tgt.evaluate(STARTQ)
                if "Pending" not in info and "Uploading" not in info and re.search(r"docid=\d", info):
                    w(">>> 两文件都完成 <<<")
                    break
            page.screenshot(path=str(paths.STATE / "probe_attach15.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
