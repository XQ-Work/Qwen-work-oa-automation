# -*- coding: utf-8 -*-
"""探附件8：过程文件上传明细表精确通道。点 addRow0 -> dump 文档上传表新行 HTML
-> 对该表内 file input 塞文件 -> 轮询行内是否出现文件名。不保存。"""
import sys
import time
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach8.txt"
TESTFILE = str(paths.DATA / "OA附件库/文档整理/整理中/采购合同（变频器）.pdf")
log = []


def w(s):
    log.append(s); print(s, flush=True)


# 定位“文档上传”表并 dump 最后一行 HTML + 行内 inputs
ROW = r"""
(function(){
  var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
  if(!tbl)return 'no 文档上传 table';
  var rows=tbl.querySelectorAll('tr');
  var last=rows[rows.length-1];
  var inputs=last?last.querySelectorAll('input'):[];
  var info=[];
  for(var j=0;j<inputs.length;j++){var e=inputs[j];
    info.push(e.type+' name='+(e.name||'')+' id='+(e.id||'')+' style='+String(e.getAttribute('style')||'').slice(0,40));}
  return 'rows='+rows.length+'\nLASTROW_HTML='+(last?last.outerHTML.replace(/\s+/g,' ').slice(0,1200):'')+'\nINPUTS:\n'+info.join('\n');
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
            if not tgt:
                w("!! 没附件帧"); return
            tgt.locator("button[id='$addbutton0$']").first.click(force=True, timeout=4000)
            time.sleep(2.5)
            w("=== 加行后 文档上传表 ===")
            w(tgt.evaluate(ROW))
            # 该表内最后一个 file input 塞文件
            tbl_fi = tgt.locator("table:has(td:text-is('文档上传')) input[type=file]")
            w("\n表内 file input count=" + str(tbl_fi.count()))
            if tbl_fi.count():
                tbl_fi.last.set_input_files(TESTFILE, timeout=6000)
                for t in range(16):
                    time.sleep(1)
                    st = tgt.evaluate(r"""(function(){
                      var tds=document.querySelectorAll('td'),tbl=null;
                      for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
                      var rows=tbl.querySelectorAll('tr');var last=rows[rows.length-1];
                      return 'rows='+rows.length+' lastText='+(last.textContent||'').replace(/\s+/g,' ').slice(0,80);})()""")
                    w(f"[{t+1}s] {st}")
                    if "采购合同" in st or ".pdf" in st.lower():
                        w(">>> 文件名已出现在行内 <<<")
                        break
            page.screenshot(path=str(paths.STATE / "probe_attach8.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
