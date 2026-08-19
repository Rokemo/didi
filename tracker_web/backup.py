# -*- coding: utf-8 -*-
"""独立备份脚本：手动或计划任务调用。"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from server import backup_db
if __name__ == "__main__":
    print("备份完成：", backup_db())
