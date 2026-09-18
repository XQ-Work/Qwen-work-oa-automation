# -*- coding: utf-8 -*-
"""探附件14：一行塞两个文件到底行不行。addRow0 一次->看 input multiple->
set 两个文件->轮询 field54276_0 值与 idnum 字段->dump 行文本。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

ROOT = paths.DATA / "OA附件库/文档整理"
F1 = str(ROOT / "整理中/文档/比质比价报告单（变频器）.pdf")
F2 = str(ROOT / "整理中/文档/供应商报价单及资质（变频器）.pdf")
OUT = paths.STATE / "probe_attach14.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


INFO = r"""
(function(){
  var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
  var fi=tbl?tbl.querySelectorAll('input[type=file]'):[];
  var out=['fileInputs='+fi.length];
  for(var j=0;j<fi.length;j++)out.push('  input'+j+' multiple='+fi[j].multiple+' accept='+(fi[j].accept||''));
  var hs=tbl?tbl.querySelectorAll('input[type=hidden]'):[];
  for(var k=0;k<hs.length;k++){var e=hs[k];if(/field54276|idnum/.test(e.name))out.push('  '+e.name+'='+e.value);}
  out.push('rowText='+(tbl?tbl.textContent.replace(/\s+/g,' ').slice(0,120):''));
  return out.join('\n');
})()
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
            w("=== 加行后 ===");  w(tgt.evaluate(INFO))
            fi = tgt.locator("table:has(td:text-is('文档上传')) input[type=file]")
            if fi.count() == 0:
                tgt.evaluate(r"""(function(){var tds=document.querySelectorAll('td'),tbl;
                  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
                  tbl.querySelectorAll('input[type=file]').forEach(function(e){e.setAttribute('data-upl','1');});})()""")
                fi = tgt.locator("input[type=file][data-upl='1']")
            w("目标 file input count=" + str(fi.count()))
            w(f"塞两个文件: {Path(F1).name} + {Path(F2).name}")
            fi.last.set_input_files([F1, F2], timeout=8000)
            for t in range(12):
                time.sleep(3)
                w(f"[{(t+1)*3}s]\n" + tgt.evaluate(INFO))
                import re
                info = tgt.evaluate(INFO)
                if re.search(r"field54276_0=\d", info):
                    w(">>> field54276_0 落定 <<<")
                    break
            page.screenshot(path=str(paths.STATE / "probe_attach14.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
