# -*- coding: utf-8 -*-
"""探附件6：搞清“过程文件上传”明细表。找含该文字的帧->列出 addrow(+)按钮->
点一次加行->对比 file input 变化并 dump 新行结构。不保存。"""
import sys
import time
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach6.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


# 找 addrow 按钮 + file input + 过程文件上传 表
SCAN = r"""
(function(){
  var out=[];
  // 含“过程文件上传”的 td
  var tds=document.querySelectorAll('td');
  var tbl=null;
  for(var i=0;i<tds.length;i++){
    if((tds[i].textContent||'').trim()==='过程文件上传'){
      var t=tds[i].closest('table'); out.push('过程文件上传@table id='+(t?t.id:''));
      tbl=t; break;
    }
  }
  // 所有 addrow 类按钮
  var btns=document.querySelectorAll('[onclick]');
  var n=0;
  for(var j=0;j<btns.length;j++){
    var oc=btns[j].getAttribute('onclick')||'';
    if(/addrow|addRow|adddetail|addrowimg|showadd/i.test(oc) && n<12){
      n++;
      out.push('ADD['+btns[j].tagName.toLowerCase()+'] id='+(btns[j].id||'')+' cls='+String(btns[j].className||'').slice(0,25)+' onclick='+oc.slice(0,90));
    }
  }
  out.push('addrowHits='+n);
  out.push('fileInput='+document.querySelectorAll('input[type=file]').length);
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
            tgt = m.find_annex_frame(page)   # “过程文件上传”在含#field-annexupload的附件帧
            if not tgt:
                w("!! 没找到附件帧"); return
            w("目标帧 url=…" + (tgt.url or "")[-60:])
            w("=== 初始扫描 ===")
            w(tgt.evaluate(SCAN))
            # 点“过程文件上传”表所在区域的 addrow：找该 table 里的 + 
            # 泛微明细表加行按钮常是 img/td onclick addrow('detail0') 或 table 的 addrowimg
            clicked = False
            for sel in ["td[onclick*='addrow']", "[onclick*='addrow']",
                        "img[onclick*='addrow']", "[onclick*='AddRow']",
                        "td[onclick*='adddetail']"]:
                try:
                    loc = tgt.locator(sel)
                    if loc.count():
                        before = tgt.evaluate(SCAN)
                        loc.first.click(force=True, timeout=3000)
                        time.sleep(2)
                        after = tgt.evaluate(SCAN)
                        w(f"\n点 {sel} (共{loc.count()}个) 后:")
                        w(after)
                        clicked = True
                        break
                except Exception as e:
                    w(f"{sel} err {str(e)[:60]}")
            if not clicked:
                w("\n没点到 addrow，dump 全帧 addrow 线索见上")
            # dump 新行的 file input 与 文档名称
            w("\n=== 加行后 file inputs & 文档名称 ===")
            fi = tgt.locator("input[type=file]")
            w("fileInput count=" + str(fi.count()))
            names = tgt.evaluate(r"""(function(){var o=[];
              document.querySelectorAll('input').forEach(function(e){
                var nm=e.name||''; if(/documentName|docname|41330|field\d+$/i.test(nm)&&e.type!='file'){o.push('in name='+nm+' id='+(e.id||''));}
              });return o.slice(0,20).join('\n');})()""")
            w(names)
            page.screenshot(path=str(paths.STATE / "probe_attach6.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
