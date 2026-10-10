# -*- coding: utf-8 -*-
"""把当前构建结果按发版命名规则复制一份，供外发。

命名规则（用户 2026-10-10 指定）：
  Hytera_Globe_450MHz_Application_Scenarios_Draft_<YYYYMMDD>_v<N>.pptx            内部版（含英文备注）
  Hytera_Globe_450MHz_Application_Scenarios_Draft_<YYYYMMDD>_v<N>_External.pptx   对外版（无备注）
  N 全局递增，不随日期重置（v1 = 2026-10-10 首次发领导的版本）。

用法：python release.py            自动取下一个版本号
      python release.py 5          指定版本号
输出目录：01_交付物/releases/
"""
import re
import shutil
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SRC_INT = REPO / '01_交付物' / '菲律宾450M方案_v8.pptx'
SRC_EXT = REPO / '01_交付物' / '菲律宾450M方案_v8_对外版.pptx'
OUT = REPO / '01_交付物' / 'releases'
STEM = 'Hytera_Globe_450MHz_Application_Scenarios_Draft'


def next_version():
    nums = [int(m.group(1)) for p in OUT.glob(STEM + '_*_v*.pptx')
            if (m := re.search(r'_v(\d+)(?:_External)?\.pptx$', p.name))]
    return max(nums, default=1) + 1     # v1 已在仓库外发出，从 v2 起


if __name__ == '__main__':
    OUT.mkdir(parents=True, exist_ok=True)
    n = int(sys.argv[1]) if len(sys.argv) > 1 else next_version()
    base = '%s_%s_v%d' % (STEM, date.today().strftime('%Y%m%d'), n)
    for src, dst in [(SRC_INT, OUT / (base + '.pptx')), (SRC_EXT, OUT / (base + '_External.pptx'))]:
        if dst.exists():
            sys.exit('已存在，不覆盖：%s' % dst)
        shutil.copy2(src, dst)
        print('released:', dst.relative_to(REPO))
