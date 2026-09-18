# -*- coding: utf-8 -*-
"""探附件4：验证 set_input_files 直塞 SWFUpload 隐藏 file input 能否挂上附件。
点“过程文件上传”页签->列 file input->set_files(合同PDF)->查附件行是否出现。不保存。"""
import sys
import time
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach4.txt"
TESTFILE = str(paths.DATA / "OA附件库/文档整理/整理中/采购合同（变频器）.pdf")
FILENAME = "采购合同（变频器）"
log = []


def w(s):
    log.append(s); print(s, flush=True)


LISTFI = r"""
(function(){
  var fi=document.querySelectorAll('input[type=file]'),out=[];
  for(var k=0;k<fi.length;k++){var e=fi[k];var st=e.getAttribute('style')||'';
    out.push('#'+k+' id='+(e.id||'')+' name='+(e.name||'')+' class='+(e.className||'')+' style='+st.slice(0,60));}
  return out.join('\n');
})()
"""
# 附件区里出现该文件名的行 / 附件计数
STATE = r"""
(function(){
  var fn=%s;
  var out=[];
  var fa=document.getElementById('field-annexupload'); if(fa) out.push('field-annexupload='+fa.value);
  var fb=document.getElementById('field-annexupload-name'); if(fb) out.push('name='+fb.value);
  var cnt=document.getElementById('field-annexupload-count'); if(cnt) out.push('count='+cnt.value);
  var hit=0;
  document.querySelectorAll('td,a,span').forEach(function(e){ if((e.textContent||'').indexOf(fn)>=0) hit++; });
  out.push('filenameHits='+hit);
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
            annfr = None
            for fr in page.frames:
                try:
                    if fr.query_selector("#field-annexupload"):
                        annfr = fr; break
                except Exception:
                    pass
            if not annfr:
                w("!! 没附件帧"); return
            try:
                annfr.locator("text=过程文件上传").first.click(timeout=3000)
            except Exception as e:
                w("点页签 err " + str(e)[:60])
            time.sleep(1.5)
            w("=== file inputs ===")
            w(annfr.evaluate(LISTFI))
            fis = annfr.locator("input[type=file]")
            n = fis.count()
            w(f"count={n}  TESTFILE={TESTFILE}")
            w("--- 初始状态 ---");  w(annfr.evaluate(STATE % json.dumps(FILENAME)))
            done = False
            for k in range(n):
                loc = fis.nth(k)
                try:
                    loc.set_input_files(TESTFILE, timeout=4000)
                except Exception as e:
                    w(f"#{k} set_input_files err {str(e)[:70]}"); continue
                time.sleep(4)
                st = annfr.evaluate(STATE % json.dumps(FILENAME))
                w(f"#{k} 塞后:\n" + st)
                if "filenameHits=" in st and not st.endswith("filenameHits=0"):
                    try:
                        hits = int(st.split("filenameHits=")[1].split("\n")[0])
                    except Exception:
                        hits = 0
                    if hits > 0:
                        w(f">>> 第{k}个 file input 成功挂上附件 <<<")
                        done = True
                        break
            page.screenshot(path=str(paths.STATE / "probe_attach4.png"), full_page=True)
            w("结果 done=" + str(done))
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
