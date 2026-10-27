

from jobspy import scrape_jobs
import subprocess
import sys
import pandas as pd

all_jobs = []

locations = ["Johannesburg, Gauteng",
             "Cape Town, Western Cape",
             "Durban, KwaZulu-Natal",
             "Pretoria, Gauteng",
             "Port Elizabeth, Eastern Cape",
             "East London, Eastern Cape",
             "Bloemfontein, Free State",]

search_terms = [
    "Junior Software Developer",
    "Junior Python Developer",
    "Junior Android Developer",
    "Android Developer",
    "Kotlin Developer",
    "Junior Java Developer",
    "Frontend Developer",
    "Junior Frontend Developer",
    "React Developer",
    "Graduate Software Developer",
    "Software Developer Intern",
    "Application Support Analyst",
    "Technical Support Specialist",
]
for term in search_terms:
    for location in locations:

        jobs = scrape_jobs(
            site_name=["indeed", "linkedin", "glassdoor"],
            search_term=term,
            location=location,
            results_wanted=20,
            hours_old=1000,
            country_indeed="South Africa",
            fetch_description=True
        )

        if not jobs.empty:
            all_jobs.append(jobs)

print(f"Found {len(jobs)} jobs")

if all_jobs:
    combined_jobs = pd.concat(all_jobs, ignore_index=True)

    # Remove duplicate listings
    combined_jobs.drop_duplicates(
        subset=["job_url"],
        inplace=True
    )

columns = ["title", "company", "location", "job_url", "description"]
available_columns = [
    column for column in columns
    if column in combined_jobs.columns
]

combined_jobs = combined_jobs[available_columns].copy()

print(combined_jobs.head())
print(f"\nFound {len(combined_jobs)} unique jobs")
combined_jobs.to_excel(
    "jobs.xlsx", index=False)

subprocess.run(
    [sys.executable, "style_jobs.py"],
    check=True
)
