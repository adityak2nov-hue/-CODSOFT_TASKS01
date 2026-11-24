"""Collect and analyze product data from Books to Scrape.

Educational example for an internship task. Use only on websites that permit
automated access; check the site's terms and robots.txt before changing BASE_URL.

Install:
    python -m pip install requests beautifulsoup4 pandas openpyxl

Run:
    python web_data_extraction.py --pages 3
    python web_data_extraction.py --pages all --output-dir results
"""

from __future__ import annotations

import argparse
import logging
import re
import time
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://books.toscrape.com/catalogue/"
START_URL = "https://books.toscrape.com/catalogue/page-1.html"
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
HEADERS = {"User-Agent": "InternshipDataCollector/1.0 (educational project)"}
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def parse_price(text: str) -> float | None:
    """Convert a price string such as '£51.77' into a numeric value."""
    match = re.search(r"\d+(?:\.\d{1,2})?", text.replace(",", ""))
    return float(match.group()) if match else None


def fetch_html(session: requests.Session, url: str, timeout: int = 20) -> str:
    response = session.get(url, headers=HEADERS, timeout=timeout)
    response.raise_for_status()
    return response.text


def extract_product(card: Any, page_url: str, session: requests.Session) -> dict[str, Any]:
    title_link = card.select_one("h3 a")
    price_node = card.select_one(".price_color")
    stock_node = card.select_one(".availability")
    rating_node = card.select_one("p.star-rating")
    image_node = card.select_one(".image_container img")

    product_url = urljoin(page_url, title_link["href"]) if title_link else None
    category = None
    if product_url:
        try:
            product_soup = BeautifulSoup(fetch_html(session, product_url), "html.parser")
            crumbs = product_soup.select(".breadcrumb li a")
            category = crumbs[-1].get_text(" ", strip=True) if len(crumbs) >= 3 else None
        except requests.RequestException as exc:
            logging.warning("Could not fetch product category (%s): %s", product_url, exc)

    rating_class = next((c for c in rating_node.get("class", []) if c in RATING_MAP), None) if rating_node else None
    image_url = urljoin(page_url, image_node.get("src", "")) if image_node else None
    return {
        "title": title_link.get("title", title_link.get_text(" ", strip=True)) if title_link else None,
        "price_gbp": parse_price(price_node.get_text(" ", strip=True)) if price_node else None,
        "rating_out_of_5": RATING_MAP.get(rating_class),
        "availability": " ".join(stock_node.stripped_strings) if stock_node else None,
        "category": category,
        "product_url": product_url,
        "image_url": image_url,
    }


def scrape(pages: int | None, delay: float) -> pd.DataFrame:
    """Scrape up to pages listing pages; None means follow pagination to end."""
    rows: list[dict[str, Any]] = []
    next_url: str | None = START_URL
    visited: set[str] = set()

    with requests.Session() as session:
        while next_url and next_url not in visited and (pages is None or len(visited) < pages):
            visited.add(next_url)
            logging.info("Reading listing page %s", len(visited))
            try:
                soup = BeautifulSoup(fetch_html(session, next_url), "html.parser")
            except requests.RequestException as exc:
                logging.error("Listing request failed (%s): %s", next_url, exc)
                break

            cards = soup.select("article.product_pod")
            for card in cards:
                try:
                    rows.append(extract_product(card, next_url, session))
                except (AttributeError, requests.RequestException) as exc:
                    logging.warning("Skipping malformed product card: %s", exc)
                time.sleep(delay)

            next_link = soup.select_one("li.next a")
            next_url = urljoin(next_url, next_link["href"]) if next_link else None

    frame = pd.DataFrame(rows)
    if not frame.empty:
        frame = frame.drop_duplicates(subset=["product_url"]).reset_index(drop=True)
        frame["price_gbp"] = pd.to_numeric(frame["price_gbp"], errors="coerce")
        frame["rating_out_of_5"] = pd.to_numeric(frame["rating_out_of_5"], errors="coerce")
        frame["availability"] = frame["availability"].str.replace(r"\s+", " ", regex=True).str.strip()
    return frame


def analyze(frame: pd.DataFrame) -> pd.DataFrame:
    """Print useful exploratory statistics and return a category summary."""
    if frame.empty:
        logging.warning("No products were collected; skipping analysis.")
        return pd.DataFrame(columns=["category", "product_count", "average_price_gbp", "average_rating"])

    print("\nDATASET OVERVIEW")
    print(f"Products collected: {len(frame)}")
    print(f"Missing values by column:\n{frame.isna().sum().to_string()}")
    print("\nNUMERIC SUMMARY")
    print(frame[["price_gbp", "rating_out_of_5"]].describe().round(2).to_string())

    summary = (
        frame.groupby("category", dropna=False)
        .agg(
            product_count=("title", "count"),
            average_price_gbp=("price_gbp", "mean"),
            average_rating=("rating_out_of_5", "mean"),
        )
        .reset_index()
        .sort_values("product_count", ascending=False)
    )
    summary[["average_price_gbp", "average_rating"]] = summary[["average_price_gbp", "average_rating"]].round(2)
    print("\nCATEGORY SUMMARY")
    print(summary.to_string(index=False))
    if frame["price_gbp"].notna().any():
        print("\nMost expensive product:")
        print(frame.loc[frame["price_gbp"].idxmax(), ["title", "price_gbp", "category"]].to_string())
        print("\nLeast expensive product:")
        print(frame.loc[frame["price_gbp"].idxmin(), ["title", "price_gbp", "category"]].to_string())
    return summary


def save_outputs(frame: pd.DataFrame, summary: pd.DataFrame, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / "products.csv"
    xlsx_path = output_dir / "products_analysis.xlsx"
    frame.to_csv(csv_path, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Products", index=False)
        summary.to_excel(writer, sheet_name="Category Summary", index=False)
    logging.info("Saved %s", csv_path.resolve())
    logging.info("Saved %s", xlsx_path.resolve())


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pages", default="3", help="Number of pages, or 'all' (default: 3)")
    parser.add_argument("--delay", type=float, default=0.5, help="Seconds between product requests (default: 0.5)")
    parser.add_argument("--output-dir", type=Path, default=Path("results"), help="Directory for CSV and Excel outputs")
    args = parser.parse_args()
    if args.delay < 0:
        parser.error("--delay must be zero or greater")
    if args.pages.lower() == "all":
        page_limit = None
    else:
        try:
            page_limit = int(args.pages)
            if page_limit < 1:
                raise ValueError
        except ValueError:
            parser.error("--pages must be a positive integer or 'all'")

    products = scrape(page_limit, args.delay)
    category_summary = analyze(products)
    save_outputs(products, category_summary, args.output_dir)


if __name__ == "__main__":
    main()
