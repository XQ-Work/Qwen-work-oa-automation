# -*- coding: utf-8 -*-
"""探付款单位信息明细表删除机制。开付款单->选付款单位(自动带出银行明细)->
dump 该明细表:每行复选框name、删除按钮onclick、序号/银行账号。只读。"""
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
sys.path.insert(0, str(BASE / "modules" / "m05_contract"))
sys.path.insert(0, str(BASE / "modules" / "m06_payment"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m05  # noqa: E402
import fill_payment as m06  # noqa: E402

OUT = paths.STATE / "probe_paydetail.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


# 找“付款单位信息”明细表：含“开户行”“银行账号”列头的那张表，dump 行结构
DUMP = r"""
(function(){
  var out=[];
  // 定位含“银行账号”表头的 table
  var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){var t=(tds[i].textContent||'').trim();
    if(t==='银行账号'){tbl=tds[i].closest('table');break;}}
  if(!tbl)return 'no 银行账号 table';
  var rows=tbl.querySelectorAll('tr');
  out.push('rows='+rows.length);
  // 复选框
  var cbs=tbl.querySelectorAll('input[type=checkbox]');
  out.push('checkbox='+cbs.length);
  for(var c=0;c<cbs.length&&c<8;c++)out.push('  cb'+c+' name='+(cbs[c].name||'')+' id='+(cbs[c].id||''));
  // 删除按钮(delrow/minus)
  var btns=tbl.parentElement.querySelectorAll('[onclick]');
  var dn=[];
  for(var j=0;j<btns.length&&dn.length<8;j++){var b=btns[j];var oc=b.getAttribute('onclick')||'';
    if(/delrow|delRow|deleteRow|deldetail|minus|delDetail/i.test(oc))dn.push(b.tagName+' onclick='+oc.slice(0,60));}
  out.push('delBtns='+dn.join(' || '));
  // 每行序号+银行账号
  for(var r=0;r<rows.length;r++){var cells=rows[r].querySelectorAll('td');
    var txt=[];for(var k=0;k<cells.length;k++){var x=(cells[k].textContent||'').trim();if(x)txt.push(x.slice(0,20));}
    if(txt.length)out.push('R'+r+': '+txt.join(' | '));}
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
            page.goto(base + m06.ADD, wait_until="domcontentloaded", timeout=40000)
            frame = m05.find_form_frame(page)
            w("frame=" + (frame.url[:50] if frame else "None"))
            # 选付款单位触发银行明细自动带出
            ok = m05.browser_select(page, frame, m06.F["付款单位名称"],
                                    "荆州浦华荆清水务有限公司", "付款单位")
            w("付款单位选中=" + str(ok))
            time.sleep(3)
            w("=== 付款单位信息明细表 ===")
            w(frame.evaluate(DUMP))
            page.screenshot(path=str(paths.STATE / "probe_paydetail.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
