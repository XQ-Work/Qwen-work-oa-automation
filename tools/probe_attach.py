# -*- coding: utf-8 -*-
"""探附件：登录并打开合同会签表单，遍历所有帧，dump 泛微「附件/上传」控件线索。
不填字段、不保存。把结果写到 state/probe_attach.txt 供分析。
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))

from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach.txt"
JS = r"""
(function(){
  var out=[];
  var fi=document.querySelectorAll('input[type=file]');
  out.push('input[type=file]='+fi.length);
  for(var k=0;k<fi.length;k++){
    var e=fi[k];
    out.push('  file#'+k+' id='+(e.id||'')+' name='+(e.name||'')+' accept='+(e.accept||'')+' vis='+(e.offsetParent?1:0));
  }
  var pat=/annex|attach|upload|addfile|showfile|selectfile|fileupload|browserfile/i;
  var all=document.querySelectorAll('a,button,span,input,div,td,embed');
  var hits=0;
  for(var i=0;i<all.length && hits<40;i++){
    var e=all[i];
    var oc=(e.getAttribute && e.getAttribute('onclick'))||'';
    var idv=(e.id||'')+' '+(e.className||'')+' '+(e.name||'')+' '+oc;
    var txt=(e.textContent||'').trim().slice(0,12);
    if((pat.test(idv)) || (/(附件|上传|添加文件)/.test(txt) && txt.length<=8)){
      hits++;
      out.push('  H<'+e.tagName.toLowerCase()+'> id='+(e.id||'')+' cls='+String(e.className||'').slice(0,30)+' txt="'+txt+'" onclick='+oc.slice(0,80));
    }
  }
  out.push('hits='+hits);
  return out.join('\n');
})()
"""

log = []


def w(s):
    log.append(s)
    print(s, flush=True)


def main():
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        try:
            base = auth.login(page)
            w("login " + base)
            page.goto(base + m.ADD, wait_until="domcontentloaded", timeout=40000)
            time.sleep(3)
            frame = m.find_form_frame(page)
            w("form_frame=" + (frame.url[:70] if frame else "None"))
            time.sleep(2)
            w("总帧数=" + str(len(page.frames)))
            for i, fr in enumerate(page.frames):
                try:
                    info = fr.evaluate(JS)
                except Exception as ex:
                    info = "eval_err " + str(ex)[:60]
                u = (fr.url or "")[-45:]
                interesting = ("input[type=file]=" in info and "input[type=file]=0" not in info) \
                    or "hits=" in info and "hits=0" not in info
                w(f"\n### frame{i} url=…{u} {'<<<' if interesting else ''}")
                if interesting or True:
                    for ln in info.splitlines():
                        w("   " + ln)
            page.screenshot(path=str(paths.STATE / "probe_attach.png"), full_page=True)
            w("\n截图 probe_attach.png")
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            try:
                ctx.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()
