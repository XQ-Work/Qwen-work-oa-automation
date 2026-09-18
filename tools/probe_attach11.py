# -*- coding: utf-8 -*-
"""探附件11：监听 MultiDocUploadByWorkflow 响应 + 轮询 field54276_0(docid) 到 40s。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach11.txt"
TESTFILE = str(paths.DATA / "OA附件库/文档整理/整理中/采购合同（变频器）.pdf")
log = []


def w(s):
    log.append(s); print(s, flush=True)


GETDOC = r"""
(function(){
  var h=document.querySelector('input[name=field54276_0]');
  var t=document.querySelector('input[name=field54275_0]');
  var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
  var body=tbl?tbl.textContent.replace(/\s+/g,' ').slice(0,60):'';
  return JSON.stringify({docid:h?h.value:'NA', name:t?t.value:'NA', tbl:body});
})()
"""


def main():
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        resp_log = []

        def on_resp(r):
            if "MultiDocUploadByWorkflow" in r.url or "DocUpload" in r.url or "upload" in r.url.lower():
                try:
                    body = r.text()[:200]
                except Exception:
                    body = "(no body)"
                resp_log.append(f"{r.status} {r.url.split('?')[0].split('/')[-1]} :: {body}")
        page.on("response", on_resp)
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
            w("set_files done")
            for t in range(8):
                time.sleep(5)
                st = tgt.evaluate(GETDOC)
                w(f"[{(t+1)*5}s] {st}")
                try:
                    import json
                    if json.loads(st).get("docid", "NULL") not in ("NULL", "", "NA"):
                        w(">>> docid 落定 <<<")
                        break
                except Exception:
                    pass
            w("\n=== 上传相关响应 ===")
            for x in resp_log:
                w("  " + x)
            page.screenshot(path=str(paths.STATE / "probe_attach11.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
