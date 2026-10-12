import requests
import pandas as pd
from datetime import datetime
import time

# ------------------------------
# PubMed API settings
# ------------------------------
EMAIL = "your_email@example.com"  # 任意のメールアドレス
NCBI_API_KEY = ""  # GitHub Secrets に設定可能

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
# PMIDから詳細情報を取得（分割して送信）
# ------------------------------
def fetch_details(pmids):
    url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    records = []
    batch_size = 200  # 一度に200件ずつ取得

    for i in range(0, len(pmids), batch_size):
        batch = pmids[i:i+batch_size]
        params = {
            "db": "pubmed",
            "id": ",".join(batch),
            "retmode": "xml",
            "tool": "pmr_full_db",
            "email": EMAIL,
        }
        if NCBI_API_KEY:
            params["api_key"] = NCBI_API_KEY

        try:
            r = requests.get(url, params=params)
            r.raise_for_status()
        except requests.exceptions.HTTPError as e:
            print(f"HTTP error: {e}. Retrying after 5 seconds...")
            time.sleep(5)
            continue

        xml = r.text
        records.extend(parse_xml_batch(xml))

        # API負荷軽減のため少し待機
        time.sleep(0.5)

    return records

# ------------------------------
# XML解析（バッチ処理用）
# ------------------------------
def parse_xml_batch(xml):
    records = []
    articles = xml.split("<PubmedArticle>")
    for art in articles[1:]:
        title = extract_tag(art, "ArticleTitle")
        abstract = extract_tag(art, "AbstractText")
        journal = extract_tag(art, "Title")
        pub_date = extract_pub_date(art)
        authors = extract_authors(art)
        pmid = extract_tag(art, "PMID")

        records.append({
            "PMID": pmid,
            "Title": title,
            "Authors": authors,
            "Journal": journal,
            "PublicationDate": pub_date,
            "Abstract": abstract,
        })
    return records

# ------------------------------
# XML解析補助関数
# ------------------------------
def extract_tag(xml, tag):
    try:
        start = xml.index(f"<{tag}>") + len(tag) + 2
        end = xml.index(f"</{tag}>", start)
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

    df_new = pd.DataFrame(new_records)
    df_new["Date Added"] = today

    df_all = pd.read_excel(EXCEL_FILE, sheet_name="PMR_All")
    existing_pmids = set(df_all["PMID"].astype(str))

    df_new_unique = df_new[~df_new["PMID"].astype(str).isin(existing_pmids)]
    print(f"New unique records: {len(df_new_unique)}")

    df_all_updated = pd.concat([df_all, df_new_unique], ignore_index=True)

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
