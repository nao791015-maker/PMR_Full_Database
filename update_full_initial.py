import requests
import pandas as pd
from datetime import datetime

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

# ------------------------------
# PubMedからPMR全件を取得
# ------------------------------
def fetch_all_pmids():
    url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
    params = {
        "db": "pubmed",
        "term": SEARCH_QUERY,
        "retmax": 20000,  # 全件取得
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
    print(f"Total PMIDs fetched: {len(pmids)}")
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

        # XML解析（簡易版）
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
# Excelに書き込む
# ------------------------------
def write_excel(records):
    df = pd.DataFrame(records)

    today = datetime.now().strftime("%Y-%m-%d")
    df["Date Added"] = today

    with pd.ExcelWriter("PMR文献データベース.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="PMR_All", index=False)
        df.to_excel(writer, sheet_name="PMR_New", index=False)

    print("Excel file created: PMR文献データベース.xlsx")

# ------------------------------
# Main
# ------------------------------
if __name__ == "__main__":
    pmids = fetch_all_pmids()
    records = fetch_details(pmids)
    write_excel(records)
