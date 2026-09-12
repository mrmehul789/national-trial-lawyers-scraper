# National Trial Lawyers Directory Scraper

A Python scraper for the public member directory at The National Trial Lawyers
(`thenationaltriallawyers.org`). It collects attorney listings filtered by
practice area and region, then exports them to an Excel spreadsheet. Useful for
legal market research and building outreach lists.

Built and maintained by DataScrape Solutions.

## What it collects

For each listing in the directory:

- name
- firm
- city
- state
- area_of_practice
- image (profile photo URL)
- profile_url
- tags (for example "Top 100" style badges)
- website_link (the attorney's own site, taken from their profile page)

## Requirements

Python 3.9 or newer, plus the packages in `requirements.txt`:

```bash
pip install -r requirements.txt
```

That installs `requests`, `beautifulsoup4`, `pandas`, and `openpyxl`.

## Setup

A `region.txt` file lives next to the script with one region slug per line.
These are the site's own region filter values, for example:

```
connecticut
california-san-diego
florida-southeast
```

A ready to use `region.txt` is included. Edit it to target the regions you want.

## Usage

Run the script:

```bash
python lawfirm4.py
```

It asks for three things:

1. How many leads to scrape (a number, or `max` for all available).
2. Areas of practice, comma separated. These are the site's practice area slugs,
   for example `civilplaintiff, criminaldefense`.
3. Regions are read automatically from `region.txt`.

The scraper loops over every practice area and region combination, paginates
through the results, removes duplicates by profile URL, then visits each profile
to pull the attorney's website link.

## Output

Results are written to `output.xlsx` with these columns:

```
name, firm, city, state, area_of_practice, image, profile_url, tags, website_link
```

## How it works

1. Loads the member directory page once to capture session cookies.
2. Calls the site's WP Grid Builder AJAX endpoint with the chosen practice area
   and region filters.
3. Parses the returned HTML cards into structured records.
4. Paginates by increasing the `_loading` count until no new records appear.
5. Visits each profile page to extract the external website link.
6. Writes everything to an Excel file.

## Notes and limits

- Practice area and region values are the site's own filter slugs. If the site
  changes them, update your inputs and `region.txt` accordingly.
- The scraper is polite by default, with short delays between requests. Keep it
  that way so you do not overload the site.
- If results come back empty, the site's page structure or endpoint may have
  changed and the CSS selectors may need updating.

## Responsible use and disclaimer

This tool reads publicly listed professional directory information. Use it
responsibly and lawfully:

- Review and respect the target site's Terms of Service and `robots.txt` before
  running it.
- The data concerns real people. Only use what you collect for legitimate,
  lawful purposes, and follow applicable privacy and anti spam laws (for example
  CAN-SPAM, GDPR, and CCPA where they apply).
- Do not republish the collected personal data.
- This project is provided for educational and lawful business research
  purposes. You are responsible for how you use it, and DataScrape Solutions
  accepts no liability for misuse.

## License

No license file is included yet. Add one (for example MIT) if you want to set
explicit usage terms for the code.
