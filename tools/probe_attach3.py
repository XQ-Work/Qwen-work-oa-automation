# -*- coding: utf-8 -*-
"""探附件3：搞清上传机制。在附件帧里搜关键 token 并打印小窗口；
点页签后重扫 input[type=file]；监听弹窗(popup)新帧。"""
import sys
import time
import json
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach3.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


TOKENS = ["uploadify", ".swf", "weaverUpload", "type=\"file\"", "type='file'",
          "showSignResourceCenter", "addannex", "annexupload", "XMLHttpRequest",
          "FormData", "action=", "selectfile", "flash", "processfile",
          "uploadFile", "initUploadify", "docDetail", "browserfile"]

SCAN = r"""
(function(){
  var toks=%s;
  var html=document.documentElement.outerHTML;
  var out=[];
  out.push('HTMLlen='+html.length);
  out.push('typeFile_now='+document.querySelectorAll('input[type=file]').length);
  for(var i=0;i<toks.length;i++){
    var t=toks[i], idx=html.toLowerCase().indexOf(t.toLowerCase());
    if(idx>=0){ out.push('TOK['+t+'] @'+idx+' :: '+html.slice(idx-20, idx+90).replace(/\s+/g,' ')); }
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
            annfr = None
            for fr in page.frames:
                try:
                    if fr.query_selector("#field-annexupload"):
                        annfr = fr; break
                except Exception:
                    pass
            if not annfr:
                w("!! 没找到附件帧"); return
            w("附件帧 url=…" + (annfr.url or "")[-70:])
            w("=== token 扫描 ===")
            w(annfr.evaluate(SCAN % json.dumps(TOKENS)))

            # 点“过程文件上传”页签，重扫 file input
            try:
                annfr.locator("text=过程文件上传").first.click(timeout=3000)
                time.sleep(2)
                w("\n点[过程文件上传]后 fileInput=" +
                  str(annfr.evaluate("()=>document.querySelectorAll('input[type=file]').length")))
                annfr.locator("text=文档上传").first.click(timeout=3000)
                time.sleep(2)
                w("点[文档上传]后 fileInput=" +
                  str(annfr.evaluate("()=>document.querySelectorAll('input[type=file]').length")))
            except Exception as e:
                w("点页签 err " + str(e)[:80])

            # 监听弹窗：点“上传”看是否 window.open
            w("\n=== 测 window.open 弹窗 ===")
            try:
                with page.expect_popup(timeout=6000) as pi:
                    annfr.locator("#field-annexupload").click(force=True, timeout=3000)
                pp = pi.value
                pp.wait_for_load_state("domcontentloaded", timeout=6000)
                w("弹出新页 url=" + pp.url)
                time.sleep(1)
                w("  新页 fileInput=" + str(pp.evaluate("()=>document.querySelectorAll('input[type=file]').length")))
                w("  新页 title=" + pp.title())
            except Exception as e:
                w("  无弹窗/或点击失败: " + str(e)[:90])
            page.screenshot(path=str(paths.STATE / "probe_attach3.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
