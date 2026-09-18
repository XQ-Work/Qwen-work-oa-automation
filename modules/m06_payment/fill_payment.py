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


# 付款信息明细表(detail0)列→字段(后缀=行号)
PI = {"费用说明": "field43271", "发票张数": "field42925",
      "费用金额": "field42926", "附件": "field43272"}
PI_ADDBTN = "$addbutton0$"   # 付款信息加行按钮


def _pi_mark_js(idx):
    return ("(function(){var o=document.querySelectorAll('input[data-upl]');"
            "for(var x=0;x<o.length;x++)o[x].removeAttribute('data-upl');"
            "var h=document.querySelector(\"input[name='field43272_%d']\");if(!h)return 0;"
            "var tr=h.closest('tr');if(!tr)return 0;var fi=tr.querySelectorAll('input[type=file]');"
            "for(var j=0;j<fi.length;j++)fi[j].setAttribute('data-upl','1');return fi.length;})()") % idx


def _pi_state_js(idx):
    return ("(function(){var h=document.querySelector(\"input[name='field43272_%d']\");"
            "var tr=h?h.closest('tr'):null;var t=tr?tr.textContent.replace(/\\s+/g,' '):'';"
            "var busy=(t.indexOf('Uploading')>=0||t.indexOf('Pending')>=0)?1:0;"
            "return (h?(h.value||''):'')+'|'+busy;})()") % idx


def fill_payinfo(frame, idx, desc, amount, count, files, step):
    """加一行“付款信息”明细并填 费用说明/发票张数/费用金额，附件传到该行附件列。"""
    try:
        frame.locator("button[id='%s']" % PI_ADDBTN).first.click(force=True, timeout=4000)
    except Exception as e:
        step("付款信息-加行", False, str(e)[:60]); return False
    time.sleep(1.5)
    try:
        frame.locator("textarea[name=%s_%d]" % (PI["费用说明"], idx)).first.fill(desc)
    except Exception as e:
        step("费用说明", False, str(e)[:60])
    try:
        frame.locator("input[name=%s_%d]" % (PI["发票张数"], idx)).first.fill(str(count))
    except Exception as e:
        step("发票张数", False, str(e)[:60])
    try:
        frame.locator("input[name=%s_%d]" % (PI["费用金额"], idx)).first.fill(str(amount))
    except Exception as e:
        step("费用金额", False, str(e)[:60])
    docid = ""
    if files:
        frame.evaluate(_pi_mark_js(idx))
        fi = frame.locator("input[type=file][data-upl='1']")
        if fi.count():
            try:
                fi.last.set_input_files([str(f) for f in files], timeout=8000)
                frame.evaluate(m05._JS_STARTQ)
                for _ in range(45):
                    time.sleep(1)
                    frame.evaluate(m05._JS_STARTQ)
                    raw = frame.evaluate(_pi_state_js(idx)) or "|"
                    d, b = (raw.split("|") + ["0"])[:2]
                    if d and d != "NULL" and b == "0":
                        docid = d
                        break
            except Exception as e:
                step("付款信息-附件", False, str(e)[:70])
    step("付款信息行%d" % (idx + 1), True,
         "金额=%s 张数=%s 附件docid=%s" % (amount, count, docid))
    return True


# 电子发票信息明细表(detail4)列→字段
EI = {"发票号码": "field43293", "发票号码对比": "field43297",
      "开票日期": "field43294", "发票金额": "field43295"}
EI_ADDBTN = "$addbutton4$"


def fill_einvoice(frame, idx, inv_no, inv_date, inv_amount, step):
    """加一行“电子发票信息”，只填 发票号码/发票号码对比(=同号防错)/开票日期/发票金额。"""
    try:
        frame.locator("button[id='%s']" % EI_ADDBTN).first.click(force=True, timeout=4000)
    except Exception as e:
        step("电子发票-加行", False, str(e)[:60]); return False
    time.sleep(1.5)
    for name, v in [(EI["发票号码"], inv_no), (EI["发票号码对比"], inv_no),
                    (EI["发票金额"], inv_amount)]:
        if not v:
            continue
        try:
            frame.locator("input[name=%s_%d]" % (name, idx)).first.fill(str(v))
        except Exception as e:
            step(f"电子发票-{name}", False, str(e)[:60])
    if inv_date:   # 开票日期是只读日期控件,JS 写值+派发 change
        js = ("(function(){var e=document.querySelector(\"input[name='field43294_%d']\");"
              "if(!e)return 0;e.value=%s;"
              "e.dispatchEvent(new Event('input',{bubbles:true}));"
              "e.dispatchEvent(new Event('change',{bubbles:true}));return 1;})()") % (idx, json.dumps(str(inv_date)))
        try:
            frame.evaluate(js)
        except Exception as e:
            step("电子发票-开票日期", False, str(e)[:60])
    step("电子发票行%d" % (idx + 1), True,
         "号码=%s 日期=%s 金额=%s" % (inv_no, inv_date, inv_amount))
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--payer", default="荆州浦华荆清水务有限公司")
    ap.add_argument("--payee", required=True)
    ap.add_argument("--amount", required=True)
    ap.add_argument("--item-type", default="货物类采购事项")
    ap.add_argument("--contract-name", default="", help="费用说明用:合同名称")
    ap.add_argument("--invoices", type=int, default=1, help="发票张数")
    ap.add_argument("--inv-no", default="", help="发票号码")
    ap.add_argument("--inv-date", default="", help="开票日期 YYYY-MM-DD")
    ap.add_argument("--inv-amount", default="", help="发票金额(价税合计)")
    ap.add_argument("--attach", nargs="*", default=[], help="付款信息附件路径")
    ap.add_argument("--save", action="store_true")
    args = ap.parse_args()

    def step(name, ok, err=""):
        print(f"[{'ok ' if ok else 'ERR'}] {name} {err}", flush=True)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(paths.PROFILE), channel="msedge", headless=False,
            args=["--start-maximized"], no_viewport=True)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        signals = []
        page.on("dialog", lambda d: (signals.append(d.message or ""), d.accept()))
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

            # 付款信息明细行：费用说明/发票张数/费用金额 + 附件
            cn = args.contract_name or ""
            desc = f"{cn}，已完成入库验收，按照合同约定付款{args.amount}元"
            atts = [Path(a) for a in args.attach if Path(a).exists()]
            fill_payinfo(frame, 0, desc, args.amount, args.invoices, atts, step)
            if args.inv_no or args.inv_date or args.inv_amount:
                fill_einvoice(frame, 0, args.inv_no, args.inv_date,
                              args.inv_amount or args.amount, step)

            time.sleep(2)
            page.screenshot(path=str(paths.STATE / "payment_filled.png"), full_page=True)
            step("整页截图", True)
            if args.save:
                del signals[:]
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
