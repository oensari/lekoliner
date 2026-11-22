
import os
import json
import time
import requests

from pathlib import Path
from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =========================================================
# 1. GENERAL SETTINGS
# =========================================================

API_URL = "https://api.openalex.org/works"
API_KEY = os.environ.get("OPENALEX_API_KEY")

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
CONFIG_FILE = BASE_DIR / "research_config.json"

OUTPUT_DIR.mkdir(exist_ok=True)


# =========================================================
# 2. SCREENING DICTIONARIES
# =========================================================

EDUCATION_TERMS = [
    "student", "students",
    "learner", "learners",
    "pupil", "pupils",
    "school", "classroom",
    "undergraduate",
    "university student",
    "physics education",
    "science education",
    "physics teaching",
    "science teaching",
    "teacher education",
    "pre-service teacher",
    "öğrenci", "öğrenciler",
    "öğretim", "öğrenme",
    "eğitim", "fizik eğitimi",
    "fen eğitimi",
    "öğretmen adayı"
]

CONCEPTUAL_TERMS = [
    "misconception", "misconceptions",
    "alternative conception",
    "alternative conceptions",
    "conceptual understanding",
    "conceptual difficulty",
    "conceptual difficulties",
    "conceptual change",
    "student understanding",
    "students' understanding",
    "student perception",
    "student reasoning",
    "learning difficulty",
    "learning difficulties",
    "conceptual knowledge",
    "mental model",
    "mental models",
    "kavram yanılgısı",
    "kavram yanılgıları",
    "alternatif kavram",
    "alternatif kavramlar",
    "kavramsal anlama",
    "kavramsal güçlük",
    "kavramsal güçlükler",
    "öğrenci düşünceleri",
    "öğrenci anlayışı",
    "öğrenme güçlüğü"
]


# =========================================================
# 3. LOAD CONFIGURATION
# =========================================================

def load_config():

    with open(CONFIG_FILE, "r", encoding="utf-8") as file:
        return json.load(file)


# =========================================================
# 4. ABSTRACT RECONSTRUCTION
# =========================================================

def reconstruct_abstract(inverted_index):

    if not inverted_index:
        return ""

    words = {}

    for word, positions in inverted_index.items():

        for position in positions:
            words[position] = word

    return " ".join(
        words[i] for i in sorted(words)
    )


# =========================================================
# 5. KEYWORD EXTRACTION
# =========================================================

def get_keywords(article):

    keywords = []

    for item in article.get("keywords", []):

        name = item.get("display_name")

        if name:
            keywords.append(name)

    return keywords


# =========================================================
# 6. SEARCHABLE TEXT
# =========================================================

def get_searchable_text(article):

    title = article.get("title") or ""

    abstract = reconstruct_abstract(
        article.get("abstract_inverted_index")
    )

    keywords = " ".join(
        get_keywords(article)
    )

    return f"{title} {abstract} {keywords}".lower()


# =========================================================
# 7. TOPIC CLASSIFICATION
# =========================================================

def classify_topics(article, research_topics):

    searchable_text = get_searchable_text(article)

    matched_topics = []
    matched_terms = {}

    for topic_name, terms in research_topics.items():

        topic_matches = []

        for term in terms:

            if term.lower() in searchable_text:
                topic_matches.append(term)

        if topic_matches:

            matched_topics.append(topic_name)

            matched_terms[topic_name] = topic_matches

    return matched_topics, matched_terms


# =========================================================
# 8. RESEARCH FOCUS CLASSIFICATION
# =========================================================

def identify_research_focus(article, research_focus):

    searchable_text = get_searchable_text(article)

    matched_focus = []

    for term in research_focus:

        if term.lower() in searchable_text:
            matched_focus.append(term)

    return matched_focus


# =========================================================
# 9. EDUCATION AND CONCEPTUAL INDICATORS
# =========================================================

def find_matching_terms(text, terms):

    matches = []

    for term in terms:

        if term.lower() in text:
            matches.append(term)

    return matches


# =========================================================
# 10. ARTICLE SCREENING
# =========================================================

def screen_article(article, research_topics, research_focus):

    searchable_text = get_searchable_text(article)

    topics, topic_terms = classify_topics(
        article,
        research_topics
    )

    education_matches = find_matching_terms(
        searchable_text,
        EDUCATION_TERMS
    )

    conceptual_matches = find_matching_terms(
        searchable_text,
        CONCEPTUAL_TERMS
    )

    focus_matches = identify_research_focus(
        article,
        research_focus
    )

    if not topics:

        status = "İlgisiz"

        reason = (
            "Araştırma konularıyla eşleşen ifade bulunamadı."
        )

    elif education_matches and (
        conceptual_matches or focus_matches
    ):

        status = "İlgili"

        reason = (
            "Araştırma konusu, eğitim ve kavramsal "
            "araştırma göstergeleri bulundu."
        )

    elif education_matches:

        status = "İncelenmeli"

        reason = (
            "Eğitim göstergesi bulundu; ancak kavramsal "
            "araştırma odağı doğrulanmalı."
        )

    else:

        status = "Alan kaynağı"

        reason = (
            "Fizik konusuyla ilişkili; ancak öğrenci "
            "veya eğitim araştırması olduğu doğrulanmadı."
        )

    return {
        "status": status,
        "reason": reason,
        "topics": topics,
        "topic_terms": topic_terms,
        "education_matches": education_matches,
        "conceptual_matches": conceptual_matches,
        "focus_matches": focus_matches
    }


# =========================================================
# 11. OPENALEX SEARCH
# =========================================================

def search_openalex(
    query,
    language,
    max_results,
    year_from,
    year_to
):

    results = []
    cursor = "*"

    per_page = min(50, max_results)

    while len(results) < max_results:

        params = {

            "search": query,

            "filter": (
                f"language:{language},"
                f"from_publication_date:{year_from}-01-01,"
                f"to_publication_date:{year_to}-12-31"
            ),

            "per-page": per_page,
            "cursor": cursor,

            "select": (
                "id,doi,title,publication_year,"
                "authorships,primary_location,"
                "type,language,abstract_inverted_index,"
                "open_access,biblio,cited_by_count,keywords"
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

            print(
                f"API error: {language} | {query} | {error}"
            )

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


# =========================================================
# 12. AUTHOR INFORMATION
# =========================================================

def get_authors(article):

    authors = []

    for item in article.get("authorships", []):

        author = item.get(
            "author", {}
        ).get("display_name")

        if author:
            authors.append(author)

    return "; ".join(authors)


# =========================================================
# 13. JOURNAL INFORMATION
# =========================================================

def get_journal(article):

    location = article.get("primary_location") or {}

    source = location.get("source") or {}

    return source.get("display_name", "")


# =========================================================
# 14. ARTICLE COLLECTION AND SCREENING
# =========================================================

def analyze_results(config):

    all_articles = {}

    languages = config["language"]

    if isinstance(languages, str):
        languages = [languages]

    queries = config["search_queries"]

    max_results = config["max_results_per_query"]

    year_from = config["publication_year_from"]

    year_to = config["publication_year_to"]

    research_topics = config["research_topics"]

    research_focus = config["research_focus"]

    for language in languages:

        for query in queries:

            print(
                f"Searching: {language} | {query}"
            )

            articles = search_openalex(
                query,
                language,
                max_results,
                year_from,
                year_to
            )

            for article in articles:

                article_id = article.get("id")

                if not article_id:
                    continue

                year = article.get("publication_year")

                if not year:
                    continue

                if not (year_from <= year <= year_to):
                    continue

                if article_id not in all_articles:

                    article["matched_queries"] = []

                    all_articles[article_id] = article

                query_label = f"{language}: {query}"

                if query_label not in all_articles[
                    article_id
                ]["matched_queries"]:

                    all_articles[
                        article_id
                    ]["matched_queries"].append(query_label)

            print(
                f"Records collected: {len(all_articles)}"
            )

            time.sleep(0.2)

    final_articles = []

    for article in all_articles.values():

        screening = screen_article(
            article,
            research_topics,
            research_focus
        )

        article["matched_topics"] = screening["topics"]

        article["matched_topic_terms"] = screening[
            "topic_terms"
        ]

        article["matched_focus"] = screening[
            "focus_matches"
        ]

        article["education_matches"] = screening[
            "education_matches"
        ]

        article["conceptual_matches"] = screening[
            "conceptual_matches"
        ]

        article["relevance_status"] = screening[
            "status"
        ]

        article["screening_reason"] = screening[
            "reason"
        ]

        final_articles.append(article)

    return final_articles


# =========================================================
# 15. EXCEL HEADERS
# =========================================================

HEADERS = [

    "Title",
    "Authors",
    "Year",
    "Journal",
    "Language",
    "Article Type",

    "Research Topics",
    "Topic Evidence",

    "Research Focus",
    "Education Indicators",
    "Conceptual Indicators",

    "Screening Status",
    "Screening Reason",

    "Abstract",
    "Keywords",

    "DOI",
    "Article URL",

    "Open Access",
    "Citations",

    "Matched Queries"
]


# =========================================================
# 16. ARTICLE ROW
# =========================================================

def create_article_row(article):

    doi = article.get("doi") or ""

    article_url = (
        f"https://doi.org/{doi}"
        if doi
        else article.get("id", "")
    )

    open_access = article.get("open_access") or {}

    topic_evidence = []

    for topic, terms in article.get(
        "matched_topic_terms", {}
    ).items():

        topic_evidence.append(
            f"{topic}: {', '.join(terms)}"
        )

    return [

        article.get("title", ""),

        get_authors(article),

        article.get("publication_year", ""),

        get_journal(article),

        article.get("language", ""),

        article.get("type", ""),

        "; ".join(article.get("matched_topics", [])),

        " | ".join(topic_evidence),

        "; ".join(article.get("matched_focus", [])),

        "; ".join(article.get("education_matches", [])),

        "; ".join(article.get("conceptual_matches", [])),

        article.get("relevance_status", ""),

        article.get("screening_reason", ""),

        reconstruct_abstract(
            article.get("abstract_inverted_index")
        ),

        "; ".join(get_keywords(article)),

        doi,

        article_url,

        open_access.get("is_oa", False),

        article.get("cited_by_count", 0),

        " | ".join(article.get("matched_queries", []))
    ]


# =========================================================
# 17. EXCEL SHEET FORMATTING
# =========================================================

def format_sheet(sheet):

    header_fill = PatternFill(
        "solid",
        fgColor="244062"
    )

    for cell in sheet[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = header_fill

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    widths = [
        45, 35, 12, 30, 12,
        18, 28, 40, 35, 35,
        35, 20, 50, 75, 35,
        35, 45, 15, 12, 55
    ]

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


# =========================================================
# 18. SUMMARY SHEET
# =========================================================

def create_summary_sheet(workbook, articles, config):

    sheet = workbook.create_sheet("Summary")

    sheet.append(["ACADEMIC RESEARCH AGENT"])
    sheet.append(["Research title", config["research_title"]])

    sheet.append([
        "Publication period",
        f"{config['publication_year_from']}–{config['publication_year_to']}"
    ])

    sheet.append([
        "Languages",
        ", ".join(config["language"])
        if isinstance(config["language"], list)
        else config["language"]
    ])

    sheet.append([])

    sheet.append(["SCREENING RESULTS", "Number of articles"])

    counts = Counter(
        article["relevance_status"]
        for article in articles
    )

    statuses = [
        "İlgili",
        "İncelenmeli",
        "Alan kaynağı",
        "İlgisiz"
    ]

    for status in statuses:

        sheet.append([
            status,
            counts.get(status, 0)
        ])

    sheet.append([])

    sheet.append([
        "RESEARCH TOPICS",
        "Number of articles"
    ])

    topic_counts = Counter()

    for article in articles:

        for topic in article.get("matched_topics", []):

            topic_counts[topic] += 1

    for topic, count in topic_counts.most_common():

        sheet.append([topic, count])

    sheet.append([])

    sheet.append([
        "Total unique records",
        len(articles)
    ])

    sheet.column_dimensions["A"].width = 35
    sheet.column_dimensions["B"].width = 75

    for cell in sheet[1]:

        cell.font = Font(
            bold=True,
            size=14,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="244062"
        )

    for row in sheet.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )


# =========================================================
# 19. CREATE SCREENING EXCEL
# =========================================================

def create_excel(articles, config):

    workbook = Workbook()

    all_sheet = workbook.active

    all_sheet.title = "All Records"

    all_sheet.append(HEADERS)

    relevant_sheet = workbook.create_sheet("Relevant")

    relevant_sheet.append(HEADERS)

    review_sheet = workbook.create_sheet("Needs Review")

    review_sheet.append(HEADERS)

    background_sheet = workbook.create_sheet("Background")

    background_sheet.append(HEADERS)

    excluded_sheet = workbook.create_sheet("Excluded")

    excluded_sheet.append(HEADERS)

    for article in articles:

        row = create_article_row(article)

        all_sheet.append(row)

        status = article["relevance_status"]

        if status == "İlgili":

            relevant_sheet.append(row)

        elif status == "İncelenmeli":

            review_sheet.append(row)

        elif status == "Alan kaynağı":

            background_sheet.append(row)

        else:

            excluded_sheet.append(row)

    for sheet in workbook.worksheets:

        if sheet.title != "Summary":

            format_sheet(sheet)

    create_summary_sheet(
        workbook,
        articles,
        config
    )

    year_from = config["publication_year_from"]

    year_to = config["publication_year_to"]

    filename = OUTPUT_DIR / (
        f"literature_screening_{year_from}_{year_to}.xlsx"
    )

    workbook.save(filename)

    print(f"Excel file created: {filename}")


# =========================================================
# 20. MAIN FUNCTION
# =========================================================

def main():

    if not API_KEY:

        raise ValueError(
            "OPENALEX_API_KEY secret is missing."
        )

    config = load_config()

    print("=" * 60)

    print("ACADEMIC RESEARCH AGENT")

    print("=" * 60)

    print(
        "Research:",
        config["research_title"]
    )

    print(
        "Publication period:",
        config["publication_year_from"],
        "-",
        config["publication_year_to"]
    )

    print("Starting literature collection...")

    articles = analyze_results(config)

    print("=" * 60)

    print("LITERATURE COLLECTION COMPLETED")

    print("=" * 60)

    print(
        "Total unique records:",
        len(articles)
    )

    counts = Counter(
        article["relevance_status"]
        for article in articles
    )

    for status, count in counts.items():

        print(f"{status}: {count}")

    create_excel(
        articles,
        config
    )

    print("=" * 60)

    print("RESEARCH SCREENING COMPLETED")

    print("=" * 60)


if __name__ == "__main__":

    main()
