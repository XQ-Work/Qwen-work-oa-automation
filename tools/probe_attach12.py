# -*- coding: utf-8 -*-
"""探附件12：set_files 后显式 startUpload()，看 docid 是否落定、上传 POST 是否发出。"""
import sys
import time
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach12.txt"
TESTFILE = str(paths.DATA / "OA附件库/文档整理/整理中/采购合同（变频器）.pdf")
log = []


def w(s):
    log.append(s); print(s, flush=True)


START = r"""
(function(){
  var n=0;
  if(window.SWFUpload&&SWFUpload.instances){for(var k in SWFUpload.instances){
    try{var u=SWFUpload.instances[k]; if(u.getStats().files_queued>0){u.startUpload();n++;}}catch(e){}}}
  return 'startUpload called on '+n+' instances';
})()
"""
GETDOC = r"""
(function(){var h=document.querySelector('input[name=field54276_0]');
  var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
  return JSON.stringify({docid:h?h.value:'NA', tbl:tbl?tbl.textContent.replace(/\s+/g,' ').slice(0,70):''});})()
"""


def main():
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        resp = []
        page.on("response", lambda r: resp.append(f"{r.status} {r.url.split('?')[0].split('/')[-1]}")
                if "MultiDocUpload" in r.url else None)
        try:
            base = auth.login(page)
            page.goto(base + m.ADD, wait_until="domcontentloaded", timeout=40000)
            time.sleep(4)
            tgt = m.find_annex_frame(page)
            tgt.locator("button[id='$addbutton0$']").first.click(force=True, timeout=4000)
            time.sleep(2)
            with page.expect_file_chooser(timeout=8000) as fi:
                tgt.locator("text=选取多个文件").last.click(force=True, timeout=4000)
            fi.value.set_files(TESTFILE)
            time.sleep(2)
            w(tgt.evaluate(START))
            for t in range(10):
                time.sleep(3)
                st = tgt.evaluate(GETDOC)
                w(f"[{(t+1)*3}s] {st}")
                try:
                    if json.loads(st).get("docid", "NULL") not in ("NULL", "", "NA"):
                        w(">>> docid 落定 <<<")
                        break
                except Exception:
                    pass
            w("上传POST: " + (", ".join(resp) or "无"))
            page.screenshot(path=str(paths.STATE / "probe_attach12.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
