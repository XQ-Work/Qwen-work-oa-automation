# -*- coding: utf-8 -*-
"""verify_attach：打开已存草稿，回读“过程文件上传”明细表是否真有那行附件。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

REQID = sys.argv[1] if len(sys.argv) > 1 else "2405438"
OUT = paths.STATE / "verify_attach.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


CHECK = r"""
(function(){
  var out=[];
  var amt=document.querySelector('input[name=field41345]');
  out.push('合同总金额field41345='+(amt?amt.value:'NA'));
  var tds=document.querySelectorAll('td'),tbl=null;
  for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='文档上传'){tbl=tds[i].closest('table');break;}}
  out.push('found_table='+(!!tbl));
  if(tbl){var rows=tbl.querySelectorAll('tr');out.push('rows='+rows.length);
    for(var r=0;r<rows.length;r++){var t=(rows[r].textContent||'').replace(/\s+/g,' ').trim();
      if(t&&t!=='序号文档名称文档上传')out.push('R'+r+': '+t.slice(0,90));}}
  var h=document.querySelector('input[name=field54276_0]');
  out.push('field54276_0='+(h?h.value:'NA'));
  var n=document.querySelector('input[name=field54275_0]');
  out.push('field54275_0(文档名称)='+(n?n.value:'NA'));
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
            u = f"{base}/workflow/request/ManageRequestNoForm.jsp?requestid={REQID}&reEdit=1&isrequest=0"
            page.goto(u, wait_until="domcontentloaded", timeout=30000)
            found = False
            for _ in range(12):          # 最多等 ~36s 让 iframe 渲染
                time.sleep(3)
                for fr in page.frames:
                    try:
                        res = fr.evaluate(CHECK)
                    except Exception:
                        continue
                    if "found_table=true" in res and "rows=" in res:
                        amt = ""
                        for ln in res.splitlines():
                            if "field41345=" in ln:
                                amt = ln
                        w("帧命中: " + (fr.url or "")[-50:])
                        w(res)
                        page.screenshot(path=str(paths.STATE / "verify_attach.png"),
                                        full_page=True)
                        found = True
                        break
                if found:
                    break
            if not found:
                w("重开未找到‘文档上传’表(可能iframe未渲染)")
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
