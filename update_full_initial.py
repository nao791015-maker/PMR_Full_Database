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

# GitHub Secrets から取得
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
        "email": EMAIL,
        "usehistory": "n"
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

        
