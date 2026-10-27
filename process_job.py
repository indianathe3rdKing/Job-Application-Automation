from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.table import TableColumn

INPUT_FILE = Path("jobs_styled.xlsx")
OUTPUT_FILE = Path("jobs_ranked.xlsx")

ROLE_KEYWORDS = {
    "Android / Kotlin": [
        "android developer",
        "junior android",
        "kotlin developer",
        "mobile developer",
    ],
    "Frontend / JavaScript": [
        "frontend developer",
        "front-end developer",
        "react developer",
        "javascript developer",
        "web developer",
    ],
    "Software Development": [
        "junior software developer",
        "software developer",
        "graduate developer",
        "graduate software",
        "python developer",
        "java developer",
        "application developer",
    ],
    "Internships / Graduate": [
        "software intern",
        "developer intern",
        "development internship",
        "graduate programme",
        "graduate program",
    ],
    "Technical Support": [
        "technical support",
        "application support",
        "it support",
        "software support",
    ],
}

SENIOR_KEYWORDS = [
    "senior",
    "lead developer",
    "principal engineer",
    "engineering manager",
    "head of engineering",
    "architect",
    "director of",
]

JUNIOR_KEYWORDS = ["junior", "entry level", "entry-level", "graduate", "intern"]
HIGH_PRIORITY_CATEGORIES = [
    "Android / Kotlin",
    "Frontend / JavaScript",
    "Software Development",
]
MEDIUM_PRIORITY_CATEGORIES = [
    "Internships / Graduate",
    "Technical Support",
]

HIGH_PRIORITY_SCORE = 50
MEDIUM_PRIORITY_SCORE = 35
SENIOR_PENALTY = 80
JUNIOR_BONUS = 20
RESULT_COLUMNS = [
    "match_score",
    "match_category",
    "match_reason",
    "application_status",
]


def score_job(title):
    title = str(title).lower()
    score = 0
    matched_category = "Other"
    reasons = []

    for category, keywords in ROLE_KEYWORDS.items():
        if any(keyword in title for keyword in keywords):
            matched_category = category
            reasons.append(f"Matches {category}")

            if category in HIGH_PRIORITY_CATEGORIES:
                score += HIGH_PRIORITY_SCORE
            elif category in MEDIUM_PRIORITY_CATEGORIES:
                score += MEDIUM_PRIORITY_SCORE

            break

    if any(keyword in title for keyword in SENIOR_KEYWORDS):
        score -= SENIOR_PENALTY
        reasons.append("May be too senior")

    if any(keyword in title for keyword in JUNIOR_KEYWORDS):
        score += JUNIOR_BONUS
        reasons.append("Entry-level indicator")

    if not reasons:
        reasons.append("No strong title match")

    score = max(0, min(score, 100))
    return score, matched_category, "; ".join(reasons)


def normalized_header(value):
    return str(value or "").strip().lower().replace(" ", "_")


def table_for_header_row(worksheet):
    for table_name in worksheet.tables:
        table = worksheet.tables[table_name]
        if range_boundaries(table.ref)[1] == 1:
            return table
    return None


def process_worksheet(worksheet):
    header_columns = {
        normalized_header(cell.value): cell.column
        for cell in worksheet[1]
        if cell.value is not None
    }
    required_columns = {"title", "company", "location", "job_url"}
    missing_columns = required_columns - header_columns.keys()

    if missing_columns:
        return False

    result_columns = {}
    for column_name in RESULT_COLUMNS:
        column = header_columns.get(column_name)
        if column is None:
            column = worksheet.max_column + 1
            worksheet.cell(row=1, column=column, value=column_name)
            header_columns[column_name] = column
        result_columns[column_name] = column

    title_column = header_columns["title"]
    url_column = header_columns["job_url"]
    max_column = worksheet.max_column
    records = []
    seen_urls = set()

    for row_number in range(2, worksheet.max_row + 1):
        cells = [
            worksheet.cell(row=row_number, column=column)
            for column in range(1, max_column + 1)
        ]
        title = cells[title_column - 1].value
        url_cell = cells[url_column - 1]
        url = url_cell.hyperlink.target if url_cell.hyperlink else url_cell.value

        if title is None or url is None:
            continue

        title = str(title).strip()
        url_key = str(url).strip()
        if not title or not url_key or url_key in seen_urls:
            continue
        seen_urls.add(url_key)

        score, category, reason = score_job(title)
        values = [cell.value for cell in cells]
        values[title_column - 1] = title
        values[result_columns["match_score"] - 1] = score
        values[result_columns["match_category"] - 1] = category
        values[result_columns["match_reason"] - 1] = reason
        values[result_columns["application_status"] - 1] = "Not applied"

        records.append(
            {
                "values": values,
                "styles": [copy(cell._style) for cell in cells],
                "hyperlinks": [
                    copy(cell.hyperlink) if cell.hyperlink else None
                    for cell in cells
                ],
                "score": score,
                "title": title,
                "company": cells[header_columns["company"] - 1].value,
                "category": category,
                "reason": reason,
            }
        )

    records.sort(key=lambda record: record["score"], reverse=True)

    header_style = copy(worksheet.cell(row=1, column=1)._style)
    row_style = copy(worksheet.cell(row=2, column=1)._style)
    header_alignment = copy(worksheet.cell(row=1, column=1).alignment)
    row_alignment = copy(worksheet.cell(row=2, column=1).alignment)
    row_border = copy(worksheet.cell(row=2, column=1).border)

    for row_number, record in enumerate(records, start=2):
        for column, value in enumerate(record["values"], start=1):
            cell = worksheet.cell(row=row_number, column=column)
            cell.value = value
            cell._style = copy(record["styles"][column - 1])
            cell.hyperlink = copy(record["hyperlinks"][column - 1])

    trailing_rows = worksheet.max_row - len(records) - 1
    if trailing_rows > 0:
        worksheet.delete_rows(len(records) + 2, trailing_rows)

    for column in result_columns.values():
        header_cell = worksheet.cell(row=1, column=column)
        header_cell._style = copy(header_style)
        header_cell.alignment = copy(header_alignment)

        for row_number in range(2, len(records) + 2):
            cell = worksheet.cell(row=row_number, column=column)
            cell._style = copy(row_style)
            cell.alignment = copy(row_alignment)
            cell.border = copy(row_border)

    table = table_for_header_row(worksheet)
    if table is not None:
        table.ref = f"A1:{worksheet.cell(row=1, column=worksheet.max_column).column_letter}{len(records) + 1}"
        if table.autoFilter is not None:
            table.autoFilter.ref = table.ref

        existing_table_columns = {
            column.name for column in table.tableColumns
        }
        for column_name, column_number in result_columns.items():
            if column_name not in existing_table_columns:
                table.tableColumns.append(
                    TableColumn(
                        id=len(table.tableColumns) + 1,
                        name=column_name,
                    )
                )

    for column_name, column in result_columns.items():
        width = {
            "match_score": 14,
            "match_category": 26,
            "match_reason": 38,
            "application_status": 18,
        }[column_name]
        worksheet.column_dimensions[
            worksheet.cell(row=1, column=column).column_letter
        ].width = width

    print(f"{worksheet.title}: ranked {len(records)} listings")
    print("Top 10 matches")
    for record in records[:10]:
        print(
            f"{record['title']} | {record['company']} | "
            f"{record['category']} | {record['score']} | {record['reason']}"
        )
    return True


def main():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Could not find {INPUT_FILE}. Run scrape_jobs.py first."
        )

    workbook = load_workbook(INPUT_FILE)
    processed_sheets = [
        worksheet
        for worksheet in workbook.worksheets
        if process_worksheet(worksheet)
    ]

    if not processed_sheets:
        raise ValueError(
            f"No worksheet in {INPUT_FILE} contains the required "
            "title, company, location, and job_url columns."
        )

    workbook.save(OUTPUT_FILE)
    total_listings = sum(
        worksheet.max_row - 1 for worksheet in processed_sheets
    )
    print(f"Ranked spreadsheet saved to: {OUTPUT_FILE.resolve()}")
    print(f"Total listings ranked: {total_listings}")


if __name__ == "__main__":
    main()
