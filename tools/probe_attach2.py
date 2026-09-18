# -*- coding: utf-8 -*-
"""探附件2：定位含 #field-annexupload 的帧，dump 附件区 HTML；再用 filechooser
事件测“点哪个触发系统选择框”，并尝试直接 set_files。不保存。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach2.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


DUMP = r"""
(function(){
  var out=[];
  var el=document.getElementById('field-annexupload');
  if(!el){return 'no #field-annexupload';}
  // 往上找 3 层容器打印
  var box=el; for(var k=0;k<6;k++){ if(box.parentElement) box=box.parentElement; }
  var html=box.outerHTML.replace(/\s+/g,' ');
  out.push('CONTAINER tag='+box.tagName+' id='+(box.id||'')+' len='+html.length);
  out.push(html.slice(0,3500));
  return out.join('\n');
})()
"""

TRIGGERS = [
    ("#field-annexupload", None),
    ("td#fsUploadProgressannexuploadSimple", None),
    ("text=上传", None),
    ("text=添加附件", None),
    ("text=过程文件上传", None),
]

TESTFILE = str(paths.DATA / "OA附件库/文档整理/整理中/采购合同（变频器）.pdf")

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
                w("!! 没找到含 #field-annexupload 的帧"); return
            w("附件帧 url=…" + (annfr.url or "")[-70:])
            w("--- 附件区 HTML ---")
            w(annfr.evaluate(DUMP))

            # 用 filechooser 事件测触发
            w("\n=== 测触发（点击后是否弹系统选择框）===")
            for sel, _ in TRIGGERS:
                try:
                    loc = annfr.locator(sel)
                    if loc.count() == 0:
                        w(f"  {sel}: 无元素"); continue
                    fired = {"got": False}

                    def on_fc(fc):
                        fired["got"] = True
                        try:
                            fc.set_files(TESTFILE)
                        except Exception as e:
                            w("   set_files err " + str(e)[:80])
                    page.on("filechooser", on_fc)
                    loc.first.click(timeout=3000)
                    time.sleep(1.5)
                    page.remove_listener("filechooser", on_fc)
                    w(f"  {sel}: count={loc.count()} -> filechooser={'FIRE✅' if fired['got'] else 'no'}")
                    if fired["got"]:
                        w("  >>> 该触发点可用！检查是否已挂上文件 <<<")
                        w("  field-annexupload value=" + str(annfr.locator("#field-annexupload").input_value()))
                        w("  field-annexupload-name value=" + str(annfr.locator("#field-annexupload-name").input_value()))
                        page.screenshot(path=str(paths.STATE / "probe_attach2.png"), full_page=True)
                        break
                except Exception as e:
                    w(f"  {sel}: err {str(e)[:70]}")
            page.screenshot(path=str(paths.STATE / "probe_attach2.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
