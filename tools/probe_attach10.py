# -*- coding: utf-8 -*-
"""探附件10：addRow0->选取多个文件 set_files->等15s->dump 文档上传整表(每行每input的name/value)。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach10.txt"
TESTFILE = str(paths.DATA / "OA附件库/文档整理/整理中/采购合同（变频器）.pdf")
log = []


def w(s):
    log.append(s); print(s, flush=True)


DUMP = r"""
(function(){
  var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
  if(!tbl)return 'no table';
  var out=[];var rows=tbl.querySelectorAll('tr');
  out.push('rows='+rows.length);
  for(var r=0;r<rows.length;r++){
    var ins=rows[r].querySelectorAll('input');var parts=[];
    for(var j=0;j<ins.length;j++){var e=ins[j];parts.push(e.type+'['+(e.name||e.id||'?')+'="'+(e.value||'').slice(0,30)+'"]');}
    var txt=(rows[r].textContent||'').replace(/\s+/g,' ').trim().slice(0,50);
    out.push('R'+r+' txt="'+txt+'" inputs: '+parts.join(' '));
  }
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
            with page.expect_file_chooser(timeout=8000) as fi:
                tgt.locator("text=选取多个文件").last.click(force=True, timeout=4000)
            fi.value.set_files(TESTFILE)
            w("set_files done, 等待上传…")
            for t in range(15):
                time.sleep(1)
            w("=== 15s 后整表 ===")
            w(tgt.evaluate(DUMP))
            page.screenshot(path=str(paths.STATE / "probe_attach10.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
