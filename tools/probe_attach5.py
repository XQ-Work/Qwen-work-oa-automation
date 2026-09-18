# -*- coding: utf-8 -*-
"""探附件5：读 SWFUpload.instances 配置认出附件上传器；对它塞文件并轮询
field-annexupload(docid) 落定时长。不保存。"""
import sys
import time
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach5.txt"
TESTFILE = str(paths.DATA / "OA附件库/文档整理/整理中/采购合同（变频器）.pdf")
log = []


def w(s):
    log.append(s); print(s, flush=True)


INST = r"""
(function(){
  if(!window.SWFUpload||!SWFUpload.instances) return 'no SWFUpload.instances';
  var out=[];
  for(var id in SWFUpload.instances){
    var s=SWFUpload.instances[id].settings||{};
    out.push('inst='+id+' post='+(s.file_post_name||'')+' url='+(s.upload_url||'')+
             ' btn='+(s.button_image_url||'')+' btnId='+(s.button_id||'')+
             ' flashC='+(s.flash_container_id||'')+' btnSel='+(s.button_action||''));
  }
  return out.join('\n');
})()
"""
DOCID = r"""
(function(){
  var fa=document.getElementById('field-annexupload');
  var cnt=document.getElementById('field-annexupload-count');
  var trs=document.querySelectorAll('#field-annexupload').length;
  return 'docid='+(fa?fa.value:'NA')+' count='+(cnt?cnt.value:'NA');
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
            except Exception:
                pass
            time.sleep(1)
            w("=== SWFUpload 实例 ===")
            w(annfr.evaluate(INST))
            w("主页面 SWFUpload 实例:")
            try:
                w(page.main_frame.evaluate(INST))
            except Exception as e:
                w(" main err " + str(e)[:60])
            # 逐 input 试探：塞文件后看 annex docid/count
            fis = annfr.locator("input[type=file]")
            n = fis.count()
            w(f"\n=== annex帧 file inputs={n} ===")
            for k in range(n):
                before = annfr.evaluate(DOCID)
                try:
                    fis.nth(k).set_input_files(TESTFILE, timeout=5000)
                except Exception as e:
                    w(f"#{k} set err {str(e)[:60]}"); continue
                got = ""
                for t in range(15):
                    time.sleep(1)
                    cur = annfr.evaluate(DOCID)
                    if cur != before:
                        got = f"第{t+1}s 变化: {cur}"
                        break
                w(f"#{k} before=[{before}] -> {got or '15s内无变化'}")
                if "docid=" in got and "docid=NA" not in got:
                    try:
                        dv = got.split("docid=")[1].split(" ")[0]
                        if dv:
                            w(f">>> #{k} 是附件上传器，docid={dv} <<<")
                    except Exception:
                        pass
            page.screenshot(path=str(paths.STATE / "probe_attach5.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
