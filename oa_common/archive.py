# -*- coding: utf-8 -*-
"""采购项目档案 项目夹定位（各模块共用）：命中既有夹复用，否则当年最大序号+1 新建"""
import datetime
import re

from oa_common import paths


def norm(s):
    return re.sub(r"[（）()【】\[\]\s\-_·、,，]", "", s or "")


def project_dir(matter_name, create=True):
    """返回项目夹 Path（不建子目录）。命中既有夹（名含事项核心词）复用，否则续号。"""
    keys = [matter_name]
    for suf in ("申购单", "申购", "采购", "服务", "及耗材", "明细"):
        for k in list(keys):
            if k.endswith(suf) and len(k) - len(suf) >= 2:
                keys.append(k[:-len(suf)])
    nk = [norm(k) for k in keys if len(norm(k)) >= 2]
    if paths.ARCHIVE.is_dir():
        for d in sorted(paths.ARCHIVE.iterdir(), reverse=True):
            if d.is_dir() and any(k in norm(d.name) for k in nk if len(k) >= 3):
                return d
    if not create:
        return None
    year = datetime.date.today().year
    mx = 0
    if paths.ARCHIVE.is_dir():
        for d in paths.ARCHIVE.iterdir():
            m = re.match(rf"^{year}-(\d{{3}})", d.name)
            if m:
                mx = max(mx, int(m.group(1)))
    return paths.ARCHIVE / f"{year}-{mx + 1:03d}-{matter_name}"
