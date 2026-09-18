# -*- coding: utf-8 -*-
"""探“付款信息”明细表。开付款单->找含“费用说明/预算科目”的表的加行按钮->加一行->
dump 该行各 input(name/type) + 附件上传 file input。只读。"""
import sys
import time
import json
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))
sys.path.insert(0, str(BASE / "modules" / "m05_contract"))
sys.path.insert(0, str(BASE / "modules" / "m06_payment"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m05  # noqa: E402
import fill_payment as m06  # noqa: E402

OUT = paths.STATE / "probe_payinfo.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


ADD_BTN = ("(function(){"
           "var bs=document.querySelectorAll(\"button[id^='$addbutton']\");var out=[];"
           "for(var i=0;i<bs.length;i++){var t=bs[i].closest('table');var h=t?t.textContent:'';"
           "var hit=/费用说明|预算科目/.test(h)?'  <<付款信息':'';"
           "out.push(bs[i].id+' onclick='+(bs[i].getAttribute('onclick')||'')+hit);}"
           "return out.join('\\n');})()")

FIND_BTN = ("(function(){var r=[];var bs=document.querySelectorAll(\"button[id^='$addbutton']\");"
            "for(var i=0;i<bs.length;i++){var t=bs[i].closest('table');"
            "if(t&&/费用说明|预算科目/.test(t.textContent))r.push(bs[i].id);}"
            "return JSON.stringify(r);})()")

ROW = ("(function(){var tds=document.querySelectorAll('td'),tbl=null;"
       "for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='费用说明'){tbl=tds[i].closest('table');break;}}"
       "if(!tbl)return 'no 付款信息 table';var out=[];"
       "var rows=tbl.querySelectorAll('tr');var hdr=null,dataRow=null;"
       "for(var r=0;r<rows.length;r++){if((rows[r].textContent||'').indexOf('费用说明')>=0)hdr=rows[r];"
       "if(rows[r].querySelector('textarea[name^=field]'))dataRow=rows[r];}"
       "var heads=[];if(hdr)hdr.querySelectorAll('td,th').forEach(function(c){var t=(c.textContent||'').trim();heads.push(t||'·');});"
       "out.push('HEADERS: '+heads.join('|'));"
       "if(dataRow){var cells=dataRow.querySelectorAll('td');var map=[];"
       "for(var c=0;c<cells.length;c++){var ins=cells[c].querySelectorAll('input,select,textarea');var nm=[];"
       "ins.forEach(function(e){if(e.name)nm.push(e.name);});map.push('['+c+']'+(heads[c]||'')+'='+(nm.join(',')||'-'));}"
       "out.push('COLS: '+map.join('  '));}"
       "return out.join('\\n');})()")


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
            w("=== 加行按钮 ===")
            w(frame.evaluate(ADD_BTN))
            ids = json.loads(frame.evaluate(FIND_BTN))
            w("\n付款信息加行按钮 id=" + str(ids))
            for bid in ids:
                frame.locator("button[id='%s']" % bid).first.click(force=True, timeout=4000)
                time.sleep(2)
                w("点 %s 后新行:" % bid)
                w(frame.evaluate(ROW))
            page.screenshot(path=str(paths.STATE / "probe_payinfo.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
