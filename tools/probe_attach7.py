# -*- coding: utf-8 -*-
"""探附件7：把 $addbuttonN$ 映射到所属明细表表头；点“文档上传”那张表的加行按钮；
dump 新行的 file input + 文档名称字段 + SWFUpload 实例。不保存。"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m  # noqa: E402

OUT = paths.STATE / "probe_attach7.txt"
log = []


def w(s):
    log.append(s); print(s, flush=True)


MAP = r"""
(function(){
  var out=[];
  var bs=document.querySelectorAll("button[id^='$addbutton']");
  for(var i=0;i<bs.length;i++){
    var b=bs[i]; var t=b.closest('table'); var hdr='';
    if(t){var rows=t.querySelectorAll('tr');
      for(var r=0;r<Math.min(rows.length,3);r++){
        var cells=rows[r].querySelectorAll('td,th'); var line=[];
        for(var c=0;c<cells.length;c++){var tx=(cells[c].textContent||'').trim(); if(tx)line.push(tx.slice(0,8));}
        if(line.length>1){hdr=line.join('|');break;}
      }}
    out.push(b.id+' onclick='+(b.getAttribute('onclick')||'')+'  表头=['+hdr.slice(0,80)+']');
  }
  return out.join('\n');
})()
"""
NEWROW = r"""
(function(){
  var out=[];
  out.push('fileInput='+document.querySelectorAll('input[type=file]').length);
  if(window.SWFUpload&&SWFUpload.instances){for(var k in SWFUpload.instances){
    var s=SWFUpload.instances[k].settings||{};
    out.push('inst '+k+' url='+(s.upload_url||'').slice(0,70)+' post='+(s.file_post_name||''));}}
  // 文档上传列附近 input
  var tds=document.querySelectorAll('td');
  for(var i=0;i<tds.length;i++){var tx=(tds[i].textContent||'').trim();
    if(/文档上传|文档名称/.test(tx)&&tx.length<8){
      var inp=tds[i].querySelectorAll('input'); var s=[];
      for(var j=0;j<inp.length;j++)s.push(inp[j].type+':'+(inp[j].name||inp[j].id||''));
      out.push('CELL['+tx+'] inputs='+s.join(','));
    }}
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
            tgt = m.find_annex_frame(page)
            if not tgt:
                w("!! 没附件帧"); return
            w("=== addbutton -> 表头映射 ===")
            w(tgt.evaluate(MAP))
            # 点表头含“文档上传”的那个 addbutton：先找它的 id
            ids = tgt.evaluate(r"""(function(){var r=[];var bs=document.querySelectorAll("button[id^='$addbutton']");
              for(var i=0;i<bs.length;i++){var t=bs[i].closest('table');var h=t?t.textContent:'';
                if(/文档上传/.test(h))r.push(bs[i].id);}return JSON.stringify(r);})()""")
            w("\n文档上传表的加行按钮 id=" + str(ids))
            import json
            for bid in json.loads(ids):
                sel = "button[id='%s']" % bid.replace("'", "\\'")
                try:
                    tgt.locator(sel).first.click(force=True, timeout=3000)
                    time.sleep(2.5)
                    w(f"点 {bid} 后:")
                    w(tgt.evaluate(NEWROW))
                except Exception as e:
                    w(f"{bid} err {str(e)[:70]}")
            page.screenshot(path=str(paths.STATE / "probe_attach7.png"), full_page=True)
        finally:
            OUT.write_text("\n".join(log), encoding="utf-8")
            time.sleep(1)
            ctx.close()


if __name__ == "__main__":
    main()
