import argparse
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://books.toscrape.com/"
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"}


def get_page(session, url, retries=3):
    for attempt in range(1, retries + 1):
        try:
            response = session.get(url, headers=HEADERS, timeout=15)
            response.raise_for_status()
            response.encoding = "utf-8"
            return BeautifulSoup(response.text, "html.parser")
        except requests.RequestException as e:
            print(f"  ошибка {e}, попытка {attempt}/{retries}")
            time.sleep(2 * attempt)
    return None


def get_categories(session):
    soup = get_page(session, BASE_URL)
    if soup is None:
        return {}
    links = soup.select("div.side_categories ul li ul li a")
    return {a.get_text(strip=True): urljoin(BASE_URL, a["href"]) for a in links}


def parse_book(card, page_url):
    link = card.select_one("h3 a")
    price_text = card.select_one("p.price_color").get_text(strip=True)
    rating_class = card.select_one("p.star-rating")["class"]
    rating = next((RATINGS[c] for c in rating_class if c in RATINGS), None)
    availability = card.select_one("p.availability").get_text(strip=True)

    return {
        "Название": link["title"],
        "Цена, £": float(price_text.replace("£", "").replace("Â", "")),
        "Наличие": "Да" if "In stock" in availability else "Нет",
        "Рейтинг": rating,
        "Ссылка": urljoin(page_url, link["href"]),
    }


def scrape(start_url, max_pages=None, delay=0.5):
    books = []
    url = start_url
    page = 0
    with requests.Session() as session:
        while url:
            page += 1
            print(f"Страница {page}: {url}")
            soup = get_page(session, url)
            if soup is None:
                print("  не удалось загрузить, останавливаюсь")
                break

            for card in soup.select("article.product_pod"):
                books.append(parse_book(card, url))

            next_link = soup.select_one("li.next a")
            url = urljoin(url, next_link["href"]) if next_link else None
            if max_pages and page >= max_pages:
                break
            time.sleep(delay)
    return books


def save_excel(books, filename):
    df = pd.DataFrame(books).drop_duplicates(subset="Ссылка")
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Книги")
        ws = writer.sheets["Книги"]
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
        widths = {"A": 60, "B": 10, "C": 10, "D": 10, "E": 70}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    return len(df)


def main():
    parser = argparse.ArgumentParser(description="Парсер каталога books.toscrape.com")
    parser.add_argument("--category", help="название категории, например Travel или Mystery")
    parser.add_argument("--pages", type=int, help="сколько страниц собрать (по умолчанию все)")
    parser.add_argument("--output", default="books.xlsx")
    parser.add_argument("--list-categories", action="store_true", help="показать список категорий")
    args = parser.parse_args()

    start_url = BASE_URL
    if args.list_categories or args.category:
        with requests.Session() as session:
            categories = get_categories(session)
        if args.list_categories:
            print("\n".join(categories))
            return
        matches = {k: v for k, v in categories.items() if k.lower() == args.category.lower()}
        if not matches:
            print(f"Категория '{args.category}' не найдена. Список: python parser.py --list-categories")
            return
        start_url = next(iter(matches.values()))

    started = time.time()
    books = scrape(start_url, max_pages=args.pages)
    if not books:
        print("Ничего не собрано")
        return

    count = save_excel(books, args.output)
    print(f"Собрано книг: {count} за {time.time() - started:.0f} сек. Файл: {args.output}")


if __name__ == "__main__":
    main()
