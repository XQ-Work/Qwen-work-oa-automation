# -*- coding: utf-8 -*-
"""m06 付款申请一条龙（071501 / workflowid=486901）v1 骨架。
复用 m05 的 sel/txt/browser_select/find_form_frame。先填锁定必填项+金额+两个放大镜。
报账类型/关联事前申请/预算科目 暂不填。提交永远人工。

用法：
  python fill_payment.py --payer 荆州浦华荆清水务有限公司 \
      --payee 上海品凡动力设备有限公司 --amount 5400.00 \
      --item-type 货物类采购事项 [--save]
"""
import argparse
import json
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE))
sys.path.insert(0, str(BASE / "modules" / "m05_contract"))
from playwright.sync_api import sync_playwright  # noqa: E402
from oa_common import paths, auth  # noqa: E402
import fill_contract as m05  # noqa: E402

WF = "486901"
ADD = ("/workflow/request/AddRequest.jsp?workflowid=" + WF +
       "&isagent=0&beagenter=0&f_weaver_belongto_userid=")

F = {
    "是否为采购类": "field494131",
    "项目公司类型": "field248130",
    "是否启迪环境相关付款": "field460630",
    "采购类事项类型": "field494132",
    "填报币种": "field42894",
    "填报金额": "field43966",
    "付款单位名称": "field249130",
    "收款单位信息": "field43298",
}


def set_money(frame, fid, val):
    """金额字段：真值框 fid 默认不可见(_format 只读显示)。JS 写值+派发 change/blur
    触发 OA checkFloat 自动格式化并算大写。"""
    js = ("(function(){var f=%s,v=%s;"
          "var e=document.querySelector('input[name=\"'+f+'\"]');"
          "if(!e)return 0;e.value=v;"
          "e.dispatchEvent(new Event('input',{bubbles:true}));"
          "e.dispatchEvent(new Event('change',{bubbles:true}));"
          "e.dispatchEvent(new Event('blur',{bubbles:true}));return 1;})()"
          ) % (json.dumps(fid), json.dumps(str(val)))
    try:
        return frame.evaluate(js)
    except Exception:
        return 0


def trim_payer_bank(frame, keep_no=2):
    """付款单位信息明细表(detail5)选中付款单位后会自动带出多行银行账号，
    只保留序号=keep_no(默认2·交通银行…0517)，勾选其余行并 deleteRow5 删除。"""
    chk = ("(function(){var keep='%s';var tds=document.querySelectorAll('td'),tbl=null;"
           "for(var i=0;i<tds.length;i++){if((tds[i].textContent||'').trim()==='银行账号')"
           "{tbl=tds[i].closest('table');break;}}"
           "if(!tbl)return 0;var n=0;tbl.querySelectorAll('tr').forEach(function(tr){"
           "var cb=tr.querySelector('input[name=check_node_5]');if(!cb)return;"
           "var seq='';tr.querySelectorAll('td').forEach(function(c){var t=(c.textContent||'').trim();"
           "if(/^[0-9]+$/.test(t)&&!seq)seq=t;});"
           "if(seq&&seq!==keep){cb.checked=true;n++;}});return n;})()") % str(keep_no)
    try:
        cnt = frame.evaluate(chk)
        if cnt:
            # 点明细表5的删除按钮(onclick=deleteRow5)，比直接 evaluate 更可靠
            try:
                frame.locator("button[onclick*='deleteRow5']").first.click(timeout=4000)
            except Exception:
                frame.evaluate("try{deleteRow5(5)}catch(e){}")
        return cnt
    except Exception:
        return -1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--payer", default="荆州浦华荆清水务有限公司")
    ap.add_argument("--payee", required=True)
    ap.add_argument("--amount", required=True)
    ap.add_argument("--item-type", default="货物类采购事项")
    ap.add_argument("--save", action="store_true")
    args = ap.parse_args()

    def step(name, ok, err=""):
        print(f"[{'ok ' if ok else 'ERR'}] {name} {err}", flush=True)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        for pg in ctx.pages[1:]:
            try:
                pg.close()
            except Exception:
                pass
        try:
            base = auth.login(page)
            step("登录", True, base)
            page.goto(base + ADD, wait_until="domcontentloaded", timeout=40000)
            frame = m05.find_form_frame(page)
            if frame is None:
                step("定位表单帧", False); raise RuntimeError("无表单帧")
            step("定位表单帧", True, frame.url[:55])

            # 必填下拉（先“是否为采购类=是”，再“采购类事项类型”，可能有级联）
            try:
                m05.sel(frame, F["是否为采购类"], "是"); step("是否为采购类=是", True)
            except Exception as e:
                step("是否为采购类", False, str(e)[:60])
            time.sleep(1.5)
            for fid, label in [("项目公司类型", "运营项目"),
                               ("是否启迪环境相关付款", "否"),
                               ("采购类事项类型", args.item_type),
                               ("填报币种", "RMB")]:
                try:
                    m05.sel(frame, F[fid], label); step(f"{fid}={label}", True)
                except Exception as e:
                    step(fid, False, str(e)[:60])
            # 金额（真值框不可见，用 set_money 写值+触发大写）
            try:
                set_money(frame, F["填报金额"], args.amount)
                time.sleep(1)
                cur = frame.locator(f"input[name={F['填报金额']}]").first.input_value()
                step(f"填报金额={args.amount}", bool(cur), f"回读={cur}")
            except Exception as e:
                step("填报金额", False, str(e)[:60])
            # 两个放大镜（行内联想，偶发不中→重试3次）
            def pick(fid, name, tag):
                for _ in range(3):
                    try:
                        if m05.browser_select(page, frame, F[fid], name, tag):
                            return True
                    except Exception:
                        pass
                    time.sleep(1)
                return False
            if pick("付款单位名称", args.payer, "付款单位"):
                step("放大镜付款单位名称", True, "val=" +
                     str(frame.locator(f"input[name={F['付款单位名称']}]").first.input_value()))
                time.sleep(2)
                n = trim_payer_bank(frame, 2)   # 只留序号2的银行账号
                step("付款单位银行明细裁剪", n >= 0, f"删{n}行,保留序号2")
            else:
                step("放大镜付款单位名称", False, args.payer)
            if pick("收款单位信息", args.payee, "收款单位"):
                step("放大镜收款单位信息", True, "val=" +
                     str(frame.locator(f"input[name={F['收款单位信息']}]").first.input_value()))
            else:
                step("放大镜收款单位信息", False, args.payee)

            time.sleep(2)
            page.screenshot(path=str(paths.STATE / "payment_filled.png"), full_page=True)
            step("整页截图", True)
            if args.save:
                signals = []
                page.on("dialog", lambda d: (signals.append(d.message or ""), d.accept()))
                page.locator("input.e8_btn_top[value=保存]").first.click(timeout=8000)
                reqid, ok = "", False
                for i in range(60):
                    time.sleep(2)
                    if any("正在保存" in s or "保存中" in s for s in signals):
                        continue
                    mm = __import__("re").search(r"requestid=(\d+)", page.url or "", __import__("re").I)
                    if mm:
                        reqid, ok = mm.group(1), True; break
                    if any("成功" in s for s in signals):
                        ok = True; break
                    if i > 3 and "AddRequest" not in (page.url or ""):
                        ok = True; break
                page.screenshot(path=str(paths.STATE / "payment_saved.png"), full_page=True)
                step("保存草稿", ok, f"requestid={reqid} 弹窗={signals[:3]}")
        finally:
            time.sleep(1)
            try:
                ctx.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()
