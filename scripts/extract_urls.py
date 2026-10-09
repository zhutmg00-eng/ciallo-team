import os
import sys
import re
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')
base = r'D:\谷雨光合计划校园挑战赛\文献资料'
files = [
    '02-谷雨品牌资料.docx',
    '03-学术文献.docx',
    '04-行业报告与法规.docx',
    '05-同类比赛获奖案例.docx'
]

url_pattern = re.compile(r'https?://[^\s\u4e00-\u9fa5\)\"\'\<\>]+')

total_urls = {}
for fname in files:
    path = os.path.join(base, fname)
    doc = Document(path)
    file_urls = []
    current_title = ""
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        if any(text.startswith(prefix) for prefix in ['1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.', '一、', '二、', '三、', '四、', '五、', '六、', '七、', '八、', '九、', '十、', '主题', '案例', 'A', 'B', 'C', 'D']):
            current_title = text[:60]
        matches = url_pattern.findall(text)
        for m in matches:
            clean_url = m.rstrip(';,，。；')
            file_urls.append((clean_url, current_title, text[:150]))
    total_urls[fname] = file_urls
    print(f"{fname}: {len(file_urls)} URLs found")

with open(r'D:\谷雨光合计划校园挑战赛\extracted_urls.txt', 'w', encoding='utf-8') as f:
    for fname, urls in total_urls.items():
        f.write(f"=== {fname} ({len(urls)}) ===\n")
        for u, title, ctx in urls:
            f.write(f"URL: {u}\nTitle: {title}\nContext: {ctx}\n---\n")

print("Done. Saved to extracted_urls.txt")
