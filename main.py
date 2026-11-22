
import os
import json
import time
import requests

from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

API_URL = "https://api.openalex.org/works"
API_KEY = os.environ.get("OPENALEX_API_KEY")

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
CONFIG_FILE = BASE_DIR / "research_config.json"

OUTPUT_DIR.mkdir(exist_ok=True)


def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as file:
        return json.load(file)


def reconstruct_abstract(inverted_index):
    if not inverted_index:
        return ""

    words = {}

    for word, positions in inverted_index.items():
        for position in positions:
            words[position] = word

    return " ".join(words[i] for i in sorted(words))


def search_openalex(query, language, max_results):
    results = []
    cursor = "*"
    per_page = min(50, max_results)

    while len(results) < max_results:
        params = {
            "search": query,
            "filter": f"language:{language}",
            "per-page": per_page,
            "cursor": cursor,
            "select": (
                "id,doi,title,publication_year,"
                "authorships,primary_location,"
                "type,language,abstract_inverted_index,"
                "open_access,biblio,cited_by_count"
            )
        }

        if API_KEY:
            params["api_key"] = API_KEY

        try:
            response = requests.get(
                API_URL,
                params=params,
                timeout=60
            )
            response.raise_for_status()
            data = response.json()

        except requests.RequestException as error:
            print(f"API error: {error}")
            break

        page_results = data.get("results", [])

        if not page_results:
            break

        results.extend(page_results)

        cursor = data.get("meta", {}).get("next_cursor")

        if not cursor:
            break

        time.sleep(0.2)

    return results[:max_results]


def get_authors(article):
    authors = []

    for item in article.get("authorships", []):
        author = item.get("author", {}).get("display_name")
        if author:
            authors.append(author)

    return "; ".join(authors)


def get_journal(article):
    location = article.get("primary_location") or {}
    source = location.get("source") or {}
    return source.get("display_name", "")


def analyze_results(config):
    all_articles = {}
    languages = config["language"]
    queries = config["search_queries"]
    max_results = config["max_results_per_query"]

    for language in languages:
        for query in queries:
            print(f"Searching: {language} | {query}")

            articles = search_openalex(
                query,
                language,
                max_results
            )

            for article in articles:
                article_id = article.get("id")

                if not article_id:
                    continue

                if article_id not in all_articles:
                    article["matched_queries"] = []
                    all_articles[article_id] = article

                all_articles[article_id]["matched_queries"].append(
                    f"{language}: {query}"
                )

    return list(all_articles.values())


def create_excel(articles):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Literature Review"

    headers = [
        "Title",
        "Authors",
        "Year",
        "Journal",
        "Language",
        "Article Type",
        "Abstract",
        "DOI",
        "Article URL",
        "Open Access",
        "Citations",
        "Matched Queries"
    ]

    sheet.append(headers)

    for article in articles:
        doi = article.get("doi") or ""
        article_url = doi or article.get("id", "")

        if doi:
            article_url = doi

        open_access = article.get("open_access") or {}

        row = [
            article.get("title", ""),
            get_authors(article),
            article.get("publication_year", ""),
            get_journal(article),
            article.get("language", ""),
            article.get("type", ""),
            reconstruct_abstract(
                article.get("abstract_inverted_index")
            ),
            doi,
            article_url,
            open_access.get("is_oa", False),
            article.get("cited_by_count", 0),
            " | ".join(article.get("matched_queries", []))
        ]

        sheet.append(row)

    # Format Excel
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            "solid",
            fgColor="244062"
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    widths = [45, 35, 12, 30, 12, 18, 70, 35, 45, 15, 12, 55]

    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[
            get_column_letter(index)
        ].width = width

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    filename = OUTPUT_DIR / "literature_review.xlsx"
    workbook.save(filename)

    print(f"Excel file created: {filename}")


def main():
    if not API_KEY:
        raise ValueError(
            "OPENALEX_API_KEY secret is missing."
        )

    config = load_config()

    print("Starting Academic Research Agent...")

    articles = analyze_results(config)

    print(f"Total unique articles: {len(articles)}")

    create_excel(articles)

    print("Research completed successfully.")


if __name__ == "__main__":
    main()
