"""一键执行:抓 CCASS -> 推 Google Sheets。

额外参数会原样透传给爬虫,例如:
    python run_all.py --end 2026-07-20     # 抓到指定日期为止
"""
import os
import subprocess
import sys

# 脚本可能被从任意目录调用(如任务计划/CI),统一切到脚本所在目录
os.chdir(os.path.dirname(os.path.abspath(__file__)))

print("=" * 50)
print("步骤 1/2: 抓取 HKEX CCASS 数据")
print("=" * 50)
r1 = subprocess.run([sys.executable, "hkex_ccass_crawler.py",
                     "--stock", "03836", "--days", "6"] + sys.argv[1:])
if r1.returncode != 0:
    sys.exit("抓取失败,中止")

print()
print("=" * 50)
print("步骤 2/2: 推送到 Google Sheets")
print("=" * 50)
r2 = subprocess.run([sys.executable, "upload_to_gsheet.py"])
if r2.returncode != 0:
    sys.exit("推送失败")

print()
print("全部完成,刷新 Google Sheets 即可查看")
