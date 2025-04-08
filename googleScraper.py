import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from keybert import KeyBERT
from pdfminer.high_level import extract_text
import tempfile
import os
import time

# 初始化 KeyBERT 模型
kw_model = KeyBERT()

# 从 PDF 提取关键词（失败时退回标题）
def extract_keywords_from_pdf_or_title(link, title, top_n=5):
    if link.lower().endswith(".pdf"):
        try:
            response = requests.get(link, timeout=10)
            if response.status_code == 200:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                    tmp_pdf.write(response.content)
                    tmp_pdf_path = tmp_pdf.name

                text = extract_text(tmp_pdf_path)
                os.remove(tmp_pdf_path)

                if text and len(text.strip()) > 50:
                    keywords = kw_model.extract_keywords(text, top_n=top_n)
                    return ', '.join([kw[0] for kw in keywords])
        except Exception as e:
            print(f"[PDF失败] {e}")

    # 如果 PDF 无法使用，就退回用标题生成
    try:
        keywords = kw_model.extract_keywords(title, top_n=top_n)
        return ', '.join([kw[0] for kw in keywords])
    except:
        return ''

# 判断是否为期刊文章
def is_journal(entry_text):
    entry_text = entry_text.lower()
    return not any(keyword in entry_text for keyword in ['conference', 'symposium', 'proceedings'])

# 提取年份
def extract_year(author_info):
    # 查找括号中的年份（通常以 "(年份)" 形式出现）
    year = None
    parts = author_info.split('-')
    for part in parts:
        if '(' in part and ')' in part:
            try:
                year = int(part.strip('()'))
                break
            except ValueError:
                continue
    return year

# 抓取 Google Scholar 数据
def scrape_scholar_articles(query, num_pages):
    articles = []
    headers = {"User-Agent": "Mozilla/5.0"}

    for page in range(num_pages):
        url = f"https://scholar.google.com/scholar?start={page*10}&q={query}&hl=en&as_sdt=0,5"
        response = requests.get(url, headers=headers)
        time.sleep(2)  # 降低请求频率
        soup = BeautifulSoup(response.text, "html.parser")
        results = soup.find_all("div", class_="gs_ri")

        for result in results:
            title_tag = result.find("h3", class_="gs_rt")
            if not title_tag:
                continue
            title = title_tag.text.strip()
            link_tag = title_tag.find("a")
            link = link_tag["href"] if link_tag else "N/A"

            author_info = result.find("div", class_="gs_a").text.strip()
            source_info = author_info.split('-')[-1].strip()

            if not is_journal(source_info):
                continue

            cited_by = "0"
            footer = result.find("div", class_="gs_fl")
            if footer:
                for a in footer.find_all("a"):
                    if "Cited by" in a.text:
                        cited_by = a.text.replace("Cited by", "").strip()
                        break

            keywords = extract_keywords_from_pdf_or_title(link, title)
            year = extract_year(author_info)

            articles.append({
                "Title": title,
                "Authors": author_info,
                "Link": link,
                "Cited By": cited_by,
                "Journal": source_info,
                "Keywords": keywords,
                "Year": year
            })

    return articles

# 保存为 Excel
def save_to_excel(articles, filename):
    df = pd.DataFrame(articles)
    df.to_excel(filename, index=False)

# 浏览文件夹
def browse_folder():
    folder_path = filedialog.askdirectory()
    entry_folder.delete(0, tk.END)
    entry_folder.insert(0, folder_path)

# 抓取按钮回调
def scrape_articles():
    query = entry_query.get()
    num_pages = int(entry_pages.get())

    label_status.config(text="正在抓取中，请稍等...")
    window.update_idletasks()

    articles = scrape_scholar_articles(query, num_pages)

    folder_path = entry_folder.get()
    filename = f"{folder_path}/scholar_articles.xlsx" if folder_path else "scholar_articles.xlsx"

    save_to_excel(articles, filename)
    label_status.config(text=f"完成！共保存 {len(articles)} 条 journal 文章。")

# GUI 界面构建
window = tk.Tk()
window.title("Google Scholar Scraper")
window.geometry("420x280")

tk.Label(window, text="关键词或文章标题:").pack()
entry_query = tk.Entry(window, width=50)
entry_query.pack()

tk.Label(window, text="要抓取的页数:").pack()
entry_pages = tk.Entry(window, width=50)
entry_pages.pack()

tk.Label(window, text="输出文件夹（可选）:").pack()
entry_folder = tk.Entry(window, width=50)
entry_folder.pack()

tk.Button(window, text="浏览文件夹", command=browse_folder).pack()
tk.Button(window, text="开始抓取", command=scrape_articles).pack()

label_status = tk.Label(window, text="")
label_status.pack()

window.mainloop()
