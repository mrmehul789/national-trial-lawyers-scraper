#!/usr/bin/env python3

import json
import time
import random
import string
import requests
from bs4 import BeautifulSoup
import pandas as pd

# ---------------------------------------------------------
# CONFIG
# ---------------------------------------------------------
BASE_URL = "https://thenationaltriallawyers.org"
PAGE_URL = BASE_URL + "/member-directory/"

AJAX_BASE = "https://thenationaltriallawyers.org/?wpgb-ajax=refresh"

OUTPUT_XLSX = "output.xlsx"

HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
    "Host": "thenationaltriallawyers.org",
    "sec-ch-ua": '"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36"
}

WPGb_BASE = {
    "is_main_query": False,
    "main_query": [],
    "permalink": "https://thenationaltriallawyers.org/member-directory/",
    "facets": [1, 2, 3, 4, 5, 10],
    "lang": "",
    "id": "oxygen-element-13",
    "is_template": "Oxygen",
    "post_id": {"0": 276, "2": 1854891},
    "paged": 1
}

# ---------------------------------------------------------
# USER INPUTS (Only target + areas; regions from file)
# ---------------------------------------------------------
def get_user_inputs():
    print("=== SCRAPER CONFIGURATION ===\n")

    # 1. How many leads?
    while True:
        target_input = input("1. How many leads to scrape? (e.g. 100 or 'max' for all): ").strip().lower()
        if target_input == "max":
            target = float('inf')
            break
        try:
            target = int(target_input)
            if target > 0:
                break
            else:
                print("Please enter a positive number or 'max'.")
        except ValueError:
            print("Invalid input. Enter a number or 'max'.")

    # 2. Areas of practice (comma-separated)
    areas_input = input("2. Areas of practice (e.g. civilplaintiff, criminaldefense or both): ").strip()
    if not areas_input:
        raise ValueError("Areas of practice cannot be empty.")
    areas_list = [a.strip() for a in areas_input.split(",") if a.strip()]

    # 3. Read regions from file
    try:
        with open("region.txt", "r", encoding="utf-8") as f:
            regions_list = [line.strip() for line in f if line.strip()]
        if not regions_list:
            raise ValueError("region.txt is empty.")
    except FileNotFoundError:
        raise FileNotFoundError("region.txt not found. Please create it with one region per line.")
    except Exception as e:
        raise ValueError(f"Error reading region.txt: {e}")

    print(f"\nTarget: {'All leads' if target == float('inf') else target}")
    print(f"Practice areas: {', '.join(areas_list)}")
    print(f"Regions ({len(regions_list)}): {', '.join(regions_list)}\n")
    return target, areas_list, regions_list

# ---------------------------------------------------------
# HTML PARSER (with area_of_practice)
# ---------------------------------------------------------
def parse_member_cards_from_html(html, area_of_practice):
    soup = BeautifulSoup(html or "", "html.parser")

    dynamic_list = (
        soup.select_one(".oxy-dynamic-list")
        or soup.find(id=lambda v: v and "dynamic_list" in v)
        or soup
    )

    cards = dynamic_list.select("div.ct-div-block.card-shadow")
    results = []

    for card in cards:
        a = card.find("a", href=True)
        profile_url = a["href"].strip() if a else ""

        img = card.find("img")
        image_url = img.get("src") if img and img.has_attr("src") else ""

        name_el = card.select_one("h6.ct-headline span, h6.ct-headline, .ct-headline span")
        name = name_el.get_text(strip=True) if name_el else ""

        city = ""
        state = ""

        city_span = card.select_one("span[id$='-116']")
        state_span = card.select_one("span[id$='-128']")

        if city_span:
            city = city_span.get_text(strip=True)
        if state_span:
            state = state_span.get_text(strip=True)

        if not (city and state):
            loc = card.select_one(".text-xs, .ct-text-block.text-xs")
            if loc:
                loc_t = loc.get_text(" ", strip=True)
                if "," in loc_t:
                    parts = [p.strip() for p in loc_t.split(",")]
                    if len(parts) >= 1:
                        city = parts[0]
                    if len(parts) >= 2 and parts[1].strip():
                        state = parts[1].split()[0]

        firm = ""
        tags = []

        text_blocks = card.select(".text-sm, .text-sm span")

        for t in text_blocks:
            txt = t.get_text(" ", strip=True)
            if not txt:
                continue

            if txt.startswith("Top"):
                tags.append(txt)
                continue

            if not firm and not txt.startswith("Top"):
                firm = txt

        tags = list(dict.fromkeys([t.strip() for t in tags if t.strip()]))

        results.append({
            "name": name,
            "firm": firm,
            "city": city,
            "state": state,
            "area_of_practice": area_of_practice,
            "image": image_url,
            "profile_url": profile_url,
            "tags": tags,
            "website_link": ""  # Placeholder
        })

    return results

# ---------------------------------------------------------
# MERGE
# ---------------------------------------------------------
def merge_results(master, new_items):
    existing_urls = {r["profile_url"] for r in master if r["profile_url"]}
    added = 0

    for it in new_items:
        url = it.get("profile_url", "")
        if url and url in existing_urls:
            continue

        master.append(it)
        added += 1

        if url:
            existing_urls.add(url)

    return added

# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------
def make_boundary():
    return "----WebKitFormBoundary" + "".join(random.choices(string.ascii_letters + string.digits, k=16))

def build_multipart_body(boundary, payload_json):
    return (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="wpgb"\r\n\r\n'
        f"{payload_json}\r\n"
        f"--{boundary}--\r\n"
    )

# ---------------------------------------------------------
# MAIN: Scrape each area + region combo
# ---------------------------------------------------------
def main():
    target, areas_list, regions_list = get_user_inputs()

    session = requests.Session()
    session.headers.update(HEADERS)

    # === STEP 1: Get cookies only ===
    print("Step 1: Fetching main page to capture cookies...")
    resp = session.get(PAGE_URL, timeout=30)
    resp.raise_for_status()
    print("Cookies captured.\n")

    all_members = []

    # === STEP 2: Loop over each AREA and each REGION ===
    total_combinations = len(areas_list) * len(regions_list)
    combo_idx = 0

    for area in areas_list:
        for region in regions_list:
            combo_idx += 1
            print(f"COMBINATION {combo_idx}/{total_combinations}: {area.upper()} + {region.upper()}")
            print("-" * 70)

            # First filtered request
            ajax_url_2 = f"{AJAX_BASE}&_areas_of_practice={area}&_regions={region}"
            print(f"First batch: {ajax_url_2}")

            payload = dict(WPGb_BASE)
            boundary = make_boundary()
            body = build_multipart_body(boundary, json.dumps(payload))

            headers_ajax = HEADERS.copy()
            headers_ajax["Content-Type"] = f"multipart/form-data; boundary={boundary}"
            headers_ajax["X-Requested-With"] = "XMLHttpRequest"
            headers_ajax["Referer"] = PAGE_URL

            try:
                ajax2 = session.post(ajax_url_2, headers=headers_ajax, data=body, timeout=30)
                ajax2.raise_for_status()
            except Exception as e:
                print(f"First AJAX failed: {e}")
                continue

            try:
                json_res = ajax2.json()
            except Exception:
                ajax_html = ajax2.text or ""
            else:
                items = json_res.get("items")
                if items and isinstance(items, list):
                    ajax_html = "".join(items)
                else:
                    ajax_html = json_res.get("posts", "") or ""

            batch = parse_member_cards_from_html(ajax_html, area)
            print(f"First batch: {len(batch)} members")

            added = merge_results(all_members, batch)
            print(f"Added {added} unique\n")

            # Pagination
            loading = 30
            iterations = 0
            max_iterations = 300
            no_new_count = 0

            while (len(all_members) < target) and iterations < max_iterations:
                ajax_url_pag = f"{AJAX_BASE}&_areas_of_practice={area}&_regions={region}&_loading={loading}"
                print(f"Loading: {loading}")

                payload = dict(WPGb_BASE)
                boundary = make_boundary()
                body = build_multipart_body(boundary, json.dumps(payload))

                headers_ajax = HEADERS.copy()
                headers_ajax["Content-Type"] = f"multipart/form-data; boundary={boundary}"
                headers_ajax["X-Requested-With"] = "XMLHttpRequest"
                headers_ajax["Referer"] = PAGE_URL

                try:
                    ajax = session.post(ajax_url_pag, headers=headers_ajax, data=body, timeout=30)
                    ajax.raise_for_status()
                except Exception as e:
                    print("AJAX error:", e)
                    time.sleep(1.0)
                    iterations += 1
                    continue

                try:
                    json_res = ajax.json()
                except Exception:
                    ajax_html = ajax.text or ""
                else:
                    items = json_res.get("items")
                    if items and isinstance(items, list):
                        ajax_html = "".join(items)
                    else:
                        ajax_html = json_res.get("posts", "") or ""

                new_members = parse_member_cards_from_html(ajax_html, area)
                print(f"Found {len(new_members)} new")

                added = merge_results(all_members, new_members)
                print(f"Added {added}")

                if added == 0:
                    no_new_count += 1
                else:
                    no_new_count = 0

                total_now = len(all_members)
                print(f"Total so far: {total_now}/{'max' if target == float('inf') else target}\n")

                if no_new_count >= 1:
                    print(f"No more for {area} + {region}\n")
                    break

                loading += 12
                iterations += 1
                time.sleep(0.45)

            if len(all_members) >= target:
                print(f"Target {target} reached. Stopping early.")
                break
        if len(all_members) >= target:
            break

    # === FETCH WEBSITE LINKS ===
    print("\nFetching website links for each profile...")
    for idx, member in enumerate(all_members, 1):
        profile_url = member.get("profile_url")
        if not profile_url:
            member["website_link"] = ""
            continue

        print(f"{idx}/{len(all_members)}: {profile_url}")

        try:
            resp = session.get(profile_url, timeout=30)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")

            website_a = soup.select_one('a.ct-link.p-3.c-trig[href][target="_blank"]')
            member["website_link"] = website_a["href"] if website_a else ""
        except Exception as e:
            print(f"Error: {e}")
            member["website_link"] = ""

        time.sleep(0.5)

    # === FINALIZE ===
    if target != float('inf') and len(all_members) > target:
        all_members = all_members[:target]
        print(f"Trimmed to {target} leads.")

    for r in all_members:
        r["tags"] = ", ".join(r.get("tags", []))

    ordered_members = []
    for r in all_members:
        ordered = {
            "name": r["name"],
            "firm": r["firm"],
            "city": r["city"],
            "state": r["state"],
            "area_of_practice": r["area_of_practice"],
            "image": r["image"],
            "profile_url": r["profile_url"],
            "tags": r["tags"],
            "website_link": r.get("website_link", "")
        }
        ordered_members.append(ordered)

    df = pd.DataFrame(ordered_members)
    df.to_excel(OUTPUT_XLSX, index=False)

    print(f"\nSUCCESS: {len(all_members)} leads from {len(regions_list)} regions saved → {OUTPUT_XLSX}")
    print("Columns: name, firm, city, state, area_of_practice, image, profile_url, tags, website_link")

# ---------------------------------------------------------
if __name__ == "__main__":
    main()