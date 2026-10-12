import requests
import pandas as pd
from datetime import datetime
import os

# ------------------------------
# PubMed API settings
# ------------------------------
EMAIL = "your_email@example.com"  # 任意のメールアドレス
NCBI_API_KEY = ""  # 必要なら設定（GitHub Actionsでは環境変数で渡す）

SEARCH_QUERY = (
    '("polymyalgia rheumatica"[MeSH Terms] OR '
    '"polymyalgia rheumatica"[Title/Abstract] OR '
    '"PMR"[Title/Abstract])'
    ' NOT ("progressive muscle relaxation"[Title/Abstract] '
    ' OR "psychomotor retardation"[Title/Abstract] '
    ' OR "plasma membrane repair"[Title/Abstract])'
)

EXCEL_FILE = "PMR文献データベース.xlsx"

# ------------------------------
# PubMedからPMR新着PMIDを取得
# ------------------------------
def fetch_pmids():
    url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
    params = {
        "db": "pubmed",
        "term": SEARCH_QUERY,
        "retmax": 20000,
        "retmode": "json",
        "tool": "pmr_full_db",
        "email": EMAIL,
    }
    if NCBI_API_KEY:
        params["api_key"] = NCBI_API_KEY

    r = requests.get(url, params=params)
    r.raise_for_status()
    data = r.json()

    pmids = data["esearchresult"]["idlist"]
    print(f"Fetched PMIDs: {len(pmids)}")
    return pmids

# ------------------------------
# PMIDから詳細情報を取得
# ------------------------------
def fetch_details(pmids):
    url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    records = []

    for pid in pmids:
        params = {
            "db": "pubmed",
            "id": pid,
            "retmode": "xml",
            "tool": "pmr_full_db",
            "email": EMAIL,
        }
        if NCBI_API_KEY:
            params["api_key"] = NCBI_API_KEY

        r = requests.get(url, params=params)
        r.raise_for_status()

        xml = r.text

        title = extract_tag(xml, "ArticleTitle")
        abstract = extract_tag(xml, "AbstractText")
        journal = extract_tag(xml, "Title")
        pub_date = extract_pub_date(xml)
        authors = extract_authors(xml)

        records.append({
            "PMID": pid,
            "Title": title,
            "Authors": authors,
            "Journal": journal,
            "PublicationDate": pub_date,
            "Abstract": abstract,
        })

    return records

# ------------------------------
# XML解析用の簡易関数
# ------------------------------
def extract_tag(xml, tag):
    try:
        start = xml.index(f"<{tag}>") + len(tag) + 2
        end = xml.index(f"</{tag}>")
        return xml[start:end]
    except:
        return ""

def extract_pub_date(xml):
    try:
        year = extract_tag(xml, "Year")
        month = extract_tag(xml, "Month")
        day = extract_tag(xml, "Day")
        return f"{year}-{month}-{day}"
    except:
        return ""

def extract_authors(xml):
    authors = []
    parts = xml.split("<Author>")
    for p in parts[1:]:
        try:
            last = extract_tag(p, "LastName")
            fore = extract_tag(p, "ForeName")
            authors.append(f"{fore} {last}")
        except:
            continue
    return ", ".join(authors)

# ------------------------------
# 差分を判定してExcelを更新
# ------------------------------
def update_excel(new_records):
    today = datetime.now().strftime("%Y-%m-%d")

    # 新着データに Date Added を付与
    df_new = pd.DataFrame(new_records)
    df_new["Date Added"] = today

    # 既存Excelを読み込み
    df_all = pd.read_excel(EXCEL_FILE, sheet_name="PMR_All")

    # 既存PMID一覧
    existing_pmids = set(df_all["PMID"].astype(str))

    # 差分抽出
    df_new_unique = df_new[~df_new["PMID"].astype(str).isin(existing_pmids)]

    print(f"New unique records: {len(df_new_unique)}")

    # PMR_All に追記
    df_all_updated = pd.concat([df_all, df_new_unique], ignore_index=True)

    # PMR_New は最新新着だけで上書き
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_all_updated.to_excel(writer, sheet_name="PMR_All", index=False)
        df_new_unique.to_excel(writer, sheet_name="PMR_New", index=False)

    print("Excel updated successfully.")

# ------------------------------
# Main
# ------------------------------
if __name__ == "__main__":
    pmids = fetch_pmids()
    records = fetch_details(pmids)
    update_excel(records)
