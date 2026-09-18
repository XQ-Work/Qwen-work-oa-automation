# -*- coding: utf-8 -*-
"""探“电子发票信息”明细表。开付款单->找含“发票号码/报销人”的表加行按钮->加一行->
dump 列-字段对应。只读。"""
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

OUT = paths.STATE / "probe_einvoice.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


FIND = ("(function(){var r=[];var bs=document.querySelectorAll(\"button[id^='$addbutton']\");"
        "for(var i=0;i<bs.length;i++){var t=bs[i].closest('table');"
        "if(t&&/发票号码|报销人/.test(t.textContent))r.push(bs[i].id+' onclick='+(bs[i].getAttribute('onclick')||''));}"
        "return JSON.stringify(r);})()")

ROW = ("(function(){var tds=document.querySelectorAll('td'),tbl=null;"
       "for(var i=0;i<tds.length;i++){var t=(tds[i].textContent||'').trim();"
       "if(t==='发票号码'){tbl=tds[i].closest('table');break;}}"
       "if(!tbl)return 'no 电子发票 table';var out=[];var rows=tbl.querySelectorAll('tr');"
       "var hdr=null,dataRow=null;"
       "for(var r=0;r<rows.length;r++){if((rows[r].textContent||'').indexOf('发票号码')>=0&&!hdr)hdr=rows[r];"
       "if(rows[r].querySelector('input[name^=field],textarea[name^=field],select[name^=field]'))dataRow=rows[r];}"
       "var heads=[];if(hdr)hdr.querySelectorAll('td,th').forEach(function(c){heads.push((c.textContent||'').trim()||'·');});"
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
        page.on("dialog", lambda d: d.accept())
        try:
            base = auth.login(page)
            page.goto(base + m06.ADD, wait_until="domcontentloaded", timeout=40000)
            frame = m05.find_form_frame(page)
            w("电子发票加行按钮=" + frame.evaluate(FIND))
            ids = json.loads(frame.evaluate(FIND.replace(
                "r.push(bs[i].id+' onclick='+(bs[i].getAttribute('onclick')||''))",
                "r.push(bs[i].id)")))
            for bid in ids:
                frame.locator("button[id='%s']" % bid).first.click(force=True, timeout=4000)
                time.sleep(2)
                w("点 %s 后:" % bid)
                w(frame.evaluate(ROW))
            page.screenshot(path=str(paths.STATE / "probe_einvoice.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
