import requests
import pandas as pd
from datetime import datetime
import time
from bs4 import BeautifulSoup

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

        time.sleep(0.5)  # API負荷軽減

    return records

# ------------------------------
# XML解析（BeautifulSoupで安定抽出）
# ------------------------------
def parse_xml_batch(xml):
    records = []
    soup = BeautifulSoup(xml, "xml")
    articles = soup.find_all("PubmedArticle")

    for art in articles:
        pmid = art.PMID.text if art.PMID else ""

        title_tag = art.find("ArticleTitle")
        title = title_tag.text if title_tag else ""

        abstract_tags = art.find_all("AbstractText")
        abstract = " ".join([t.text for t in abstract_tags]) if abstract_tags else ""

        journal_tag = art.find("Title")
        journal = journal_tag.text if journal_tag else ""

        pub_date = extract_pub_date(art)
        authors = extract_authors(art)

        records.append({
            "PMID": pmid,
            "Title": title,
            "Authors": authors,
            "Journal": journal,
            "PublicationDate": pub_date,
            "Abstract": abstract
        })

    return records

# ------------------------------
# XML解析補助関数
# ------------------------------
def extract_pub_date(art):
    try:
        year = art.find("Year").text if art.find("Year") else ""
        month = art.find("Month").text if art.find("Month") else ""
        day = art.find("Day").text if art.find("Day") else ""
        return f"{year}-{month}-{day}"
    except:
        return ""

def extract_authors(art):
    authors = []
    for a in art.find_all("Author"):
        last = a.find("LastName").text if a.find("LastName") else ""
        fore = a.find("ForeName").text if a.find("ForeName") else ""
        if fore or last:
            authors.append(f"{fore} {last}")
    return ", ".join(authors)

# ------------------------------
# 差分を判定してExcelを更新
# ------------------------------
def update_excel(new_records):
    today = datetime.now().strftime("%Y-%m-%d")

    # 列順を固定
    columns = ["PMID", "Title", "Authors", "Journal", "PublicationDate", "Abstract", "Date Added"]

    df_new = pd.DataFrame(new_records)
    df_new["Date Added"] = today
    df_new = df_new[columns]  # 列順固定

    # ブランクシート対応：空の場合は新規作成
    try:
        df_all = pd.read_excel(EXCEL_FILE, sheet_name="PMR_All")
    except:
        df_all = pd.DataFrame(columns=columns)

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
