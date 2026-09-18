# -*- coding: utf-8 -*-
"""探附件9：过程文件上传正确通道。addRow0 加行 -> expect_file_chooser 点“选取多个文件”
-> set_files(合同PDF) -> 看行内是否出现文件名、文档名称是否需手填。不保存。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach9.txt"
TESTFILE = str(paths.DATA / "OA附件库/文档整理/整理中/采购合同（变频器）.pdf")
log = []


def w(s):
    log.append(s); print(s, flush=True)


ROW = r"""
(function(){
  var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
  if(!tbl)return 'no table';
  var rows=tbl.querySelectorAll('tr');
  var out=['rows='+rows.length];
  var last=rows[rows.length-1];
  out.push('lastText='+(last.textContent||'').replace(/\s+/g,' ').slice(0,90));
  // 文档名称输入框(name 里含 field) 在表内
  var ins=tbl.querySelectorAll('input');var s=[];
  for(var j=0;j<ins.length;j++){var e=ins[j];if(e.type!='file')s.push(e.type+':'+(e.name||e.id||''));}
  out.push('nameInputs='+s.join(','));
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
            if not tgt:
                w("!! 没附件帧"); return
            tgt.locator("button[id='$addbutton0$']").first.click(force=True, timeout=4000)
            time.sleep(2)
            w("=== 加行后 ===");  w(tgt.evaluate(ROW))
            btn = tgt.locator("text=选取多个文件")
            w("\n‘选取多个文件’按钮 count=" + str(btn.count()))
            if btn.count():
                try:
                    with page.expect_file_chooser(timeout=6000) as fi:
                        btn.last.click(force=True, timeout=4000)
                    fc = fi.value
                    fc.set_files(TESTFILE)
                    w(">>> file_chooser 触发并 set_files 成功 <<<")
                except Exception as e:
                    w("expect_file_chooser 失败: " + str(e)[:120])
                for t in range(16):
                    time.sleep(1)
                    st = tgt.evaluate(ROW)
                    w(f"[{t+1}s] " + st.replace("\n", " | "))
                    if "采购合同" in st or ".pdf" in st.lower():
                        w(">>> 行内出现文件名 <<<")
                        break
            page.screenshot(path=str(paths.STATE / "probe_attach9.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
