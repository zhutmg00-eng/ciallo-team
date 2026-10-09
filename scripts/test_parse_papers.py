import os
import sys
import re
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')
base = r'D:\谷雨光合计划校园挑战赛\文献资料'

def extract_structured_entries():
    # 1. Academic papers (03-学术文献.docx)
    doc3 = Document(os.path.join(base, '03-学术文献.docx'))
    papers = []
    for t_idx, t in enumerate(doc3.tables):
        for r_idx, r in enumerate(t.rows[1:]):
            c = [cell.text.strip().replace('\n', ' ') for cell in r.cells]
            if len(c) >= 8:
                num, title, author, year, journal, link, concl, usage = c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7]
                papers.append({
                    'table': t_idx + 1,
                    'num': num,
                    'title': title,
                    'author': author,
                    'year': year,
                    'journal': journal,
                    'link': link,
                    'summary': concl,
                    'usage': usage
                })
    print(f"Loaded {len(papers)} academic papers from tables.")
    return papers

papers = extract_structured_entries()
for p in papers[:5]:
    print(p['num'], p['title'][:40], '-->', p['link'])
