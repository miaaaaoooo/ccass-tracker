"""诊断脚本:抓一次 HKEX,看真实返回的 HTML 结构。"""
import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}
URL = "https://www3.hkexnews.hk/sdw/search/searchsdw.aspx"

s = requests.Session()
s.headers.update(HEADERS)

# 1) GET 拿表单
r = s.get(URL, timeout=20)
soup = BeautifulSoup(r.text, "lxml")
state = {i["name"]: i.get("value", "") for i in soup.select("input[type=hidden]") if i.get("name")}
print(f"[GET] {r.status_code}, 隐藏字段 {len(state)} 个: {list(state.keys())[:5]}...")

# 2) 列出表单里所有 input 的真实 name(很关键)
print("\n[表单所有 input name]:")
for inp in soup.select("input, select"):
    n = inp.get("name")
    if n and not n.startswith("__"):
        print(f"  {inp.name}  name={n}  id={inp.get('id')}")

# 3) POST 查询
payload = {**state,
    "__EVENTTARGET": "btnSearch", "__EVENTARGUMENT": "",
    "txtShareholdingDate": "2026/05/27",
    "txtStockCode": "03836",
    "txtStockName": "", "txtParticipantID": "", "txtParticipantName": "",
    "sortBy": "shareholding", "sortDirection": "desc",
    "btnSearch.x": "1", "btnSearch.y": "1"}
r2 = s.post(URL, data=payload, timeout=30)
print(f"\n[POST] {r2.status_code}, 响应长度 {len(r2.text)} 字符")

# 4) 把响应存到文件
with open("debug_response.html", "w", encoding="utf-8") as f:
    f.write(r2.text)
print("已保存 debug_response.html")

# 5) 列出响应里所有 <table>,看哪个是数据表
soup2 = BeautifulSoup(r2.text, "lxml")
print("\n[响应里的所有 table]:")
for i, t in enumerate(soup2.find_all("table")):
    rows = t.find_all("tr")
    print(f"  table#{i}  class={t.get('class')}  id={t.get('id')}  行数={len(rows)}")
    if rows and len(rows) > 1:
        sample = [td.get_text(strip=True)[:30] for td in rows[1].find_all(["td","th"])][:5]
        print(f"      第2行示例: {sample}")