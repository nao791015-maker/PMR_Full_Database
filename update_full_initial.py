import os
import time
import requests
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup

# ------------------------------
# Settings
# ------------------------------

EMAIL = "your_email@example.com"

# GitHub Secretsから取得
NCBI_API_KEY = os.getenv("NCBI_API_KEY")

EXCEL_FILE = "PMR文献データベース.xlsx"

SEARCH_QUERY = (
    '("polymyalgia rheumatica"[MeSH Terms] OR '
    '"polymyalgia rheumatica"[Title/Abstract] OR '
    '"PMR"[Title/Abstract])'
    ' NOT ("progressive muscle relaxation"[Title/Abstract] '
    ' OR "psychomotor retardation"[Title/Abstract] '
    ' OR "plasma membrane repair"[Title/Abstract])'
)

COLUMNS = [
    "PMID",
    "Title",
    "Authors",
    "Journal",
    "PublicationDate",
    "Abstract",
    "Date Added"
]

# ------------------------------
# PubMed検索
# ------------------------------

def fetch_all_pmids():

    url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"

    params = {
        "db": "pubmed",
        "term": SEARCH_QUERY,
        "retmax": 20000,
        "retmode": "json",
        "email": EMAIL
    }

    if NCBI_API_KEY:
        params["api_key"] = NCBI_API_KEY

    response = requests.get(url, params=params, timeout=60)
    response.raise_for_status()

    data = response.json()

    pmids = data["esearchresult"]["idlist"]

    print(f"Total PMIDs found: {len(pmids)}")

    return pmids

# ------------------------------
# 詳細情報取得
# ------------------------------

def fetch_details(pmids):

    url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"

    all_records = []

    batch_size = 200

    for i in range(0, len(pmids), batch_size):

        batch = pmids[i:i + batch_size]

        print(
            f"Processing {i + 1} - {min(i + batch_size, len(pmids))}"
        )

        params = {
            "db": "pubmed",
            "id": ",".join(batch),
            "retmode": "xml"
        }

        if NCBI_API_KEY:
            params["api_key"] = NCBI_API_KEY

        try:

            response = requests.get(
                url,
                params=params,
                timeout=60
            )

            response.raise_for_status()

        except Exception as e:

            print(
                f"Error fetching batch: {e}"
            )

            time.sleep(5)

            continue

        all_records.extend(
            parse_xml(response.text)
        )

        time.sleep(0.5)

    return all_records

# ------------------------------
# XML解析
# ------------------------------

def parse_xml(xml):

    soup = BeautifulSoup(xml, "xml")

    records = []

    for article in soup.find_all("PubmedArticle"):

        try:

            pmid = (
                article.PMID.text
                if article.PMID
                else ""
            )

            title = (
                article.ArticleTitle.get_text(
                    " ",
                    strip=True
                )
                if article.ArticleTitle
                else ""
            )

            abstract = " ".join(
                [
                    tag.get_text(
                        " ",
                        strip=True
                    )
                    for tag in article.find_all(
                        "AbstractText"
                    )
                ]
            )

            journal = ""

            journal_node = article.find(
                "Journal"
            )

            if (
                journal_node
                and journal_node.find("Title")
            ):
                journal = journal_node.find(
                    "Title"
                ).text

            authors = []

            for author in article.find_all(
                "Author"
            ):

                fore = (
                    author.ForeName.text
                    if author.ForeName
                    else ""
                )

                last = (
                    author.LastName.text
                    if author.LastName
                    else ""
                )

                name = f"{fore} {last}".strip()

                if name:
                    authors.append(name)

            publication_date = ""

            pubdate = article.find("PubDate")

            if pubdate:

                year = pubdate.find("Year")
                month = pubdate.find("Month")
                day = pubdate.find("Day")
                medline = pubdate.find("MedlineDate")

                if year:

                    publication_date = year.text

                    if month:
                        publication_date += f"-{month.text}"

                    if day:
                        publication_date += f"-{day.text}"

                elif medline:

                    publication_date = medline.text

            records.append(
                {
                    "PMID": pmid,
                    "Title": title,
                    "Authors": "; ".join(authors),
                    "Journal": journal,
                    "PublicationDate": publication_date,
                    "Abstract": abstract
                }
            )

        except Exception as e:

            print(
                f"Error parsing article: {e}"
            )

    return records

# ------------------------------
# Excel出力
# ------------------------------

def save_initial_excel(records):

    today = datetime.today().strftime(
        "%Y-%m-%d"
    )

    df_all = pd.DataFrame(records)

    df_all["Date Added"] = today

    df_all = df_all.reindex(
        columns=COLUMNS
    )

    # PMID重複除去
    df_all = df_all.drop_duplicates(
        subset=["PMID"]
    )

    # 初回は PMR_New にも同じ内容を出力
    df_new = df_all.copy()

    with pd.ExcelWriter(
        EXCEL_FILE,
        engine="openpyxl"
    ) as writer:

        df_all.to_excel(
            writer,
            sheet_name="PMR_All",
            index=False
        )

        df_new.to_excel(
            writer,
            sheet_name="PMR_New",
            index=False
        )

    print(
        f"PMR_All records : {len(df_all)}"
    )

    print(
        f"PMR_New records : {len(df_new)}"
    )

# ------------------------------
# Main
# ------------------------------

if __name__ == "__main__":

    pmids = fetch_all_pmids()

    records = fetch_details(
        pmids
    )

    save_initial_excel(
        records
    )
