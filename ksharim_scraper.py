"""
קשרים בריבוע - Scraper יומי אוטומטי
מאגר חידות לצורך למידת AI
"""

import asyncio
import json
import re
from datetime import date, datetime
from pathlib import Path

from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

URL = "https://ksharim-baribua.com/"
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

JSON_FILE  = DATA_DIR / "puzzles_database.json"
EXCEL_FILE = DATA_DIR / "puzzles_database.xlsx"
LOG_FILE   = DATA_DIR / "scrape_log.txt"


async def fetch_puzzle_data() -> dict:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            locale="he-IL",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        page = await context.new_page()

        api_responses = []

        async def handle_response(response):
            if response.status == 200 and any(k in response.url.lower() for k in ["puzzle","game","word","api"]):
                try:
                    body = await response.json()
                    api_responses.append({"url": response.url, "data": body})
                except Exception:
                    pass

        page.on("response", handle_response)
        await page.goto(URL, wait_until="networkidle", timeout=30000)
        await page.wait_for_timeout(4000)

        raw_text = await page.inner_text("body")

        words = await page.evaluate("""
            () => {
                const results = [];
                const selectors = [
                    '[class*="word"]','[class*="Word"]','[class*="cell"]','[class*="Cell"]',
                    '[class*="tile"]','[class*="Tile"]','[class*="letter"]','[class*="Letter"]',
                    '[class*="clue"]','[class*="Clue"]','[class*="answer"]','[class*="Answer"]',
                    '[class*="group"]','[class*="Group"]','[class*="category"]','[class*="Category"]',
                    'button','span[data-word]','div[data-word]'
                ];
                for (const sel of selectors) {
                    document.querySelectorAll(sel).forEach(el => {
                        const txt = el.innerText?.trim();
                        if (txt && txt.length > 0 && txt.length < 50)
                            results.push({ selector: sel, text: txt });
                    });
                }
                return results;
            }
        """)

        await browser.close()

        return {
            "raw_text":      raw_text,
            "dom_elements":  words,
            "api_responses": api_responses,
        }


def parse_puzzle(raw: dict) -> dict:
    today = date.today().isoformat()
    puzzle = {
        "date":       today,
        "fetched_at": datetime.now().isoformat(),
        "categories": [],
        "all_words":  [],
        "raw_snippet": raw["raw_text"][:2000],
    }

    for resp in raw.get("api_responses", []):
        data = resp["data"]
        if isinstance(data, dict):
            for key in ["categories","groups","words","puzzle","game"]:
                if key in data:
                    puzzle["categories"].append({"source": resp["url"], "key": key, "value": data[key]})

    dom = raw.get("dom_elements", [])
    if dom:
        puzzle["all_words"] = list({d["text"] for d in dom if len(d["text"]) > 1})

    hebrew_words = re.findall(r'[\u05d0-\u05ea]{2,}', raw["raw_text"])
    puzzle["hebrew_words_found"] = list(dict.fromkeys(hebrew_words))[:100]

    return puzzle


def save_to_json(puzzle: dict):
    if JSON_FILE.exists():
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            db = json.load(f)
    else:
        db = {"puzzles": [], "total": 0, "created": date.today().isoformat()}

    existing = [i for i, p in enumerate(db["puzzles"]) if p["date"] == puzzle["date"]]
    if existing:
        db["puzzles"][existing[0]] = puzzle
    else:
        db["puzzles"].append(puzzle)

    db["total"] = len(db["puzzles"])
    db["last_updated"] = datetime.now().isoformat()

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)
    print(f"✅ JSON נשמר")


def save_to_excel(puzzle: dict):
    HEADER_COLOR = "2E4057"
    ROW_ODD      = "F0F4F8"
    ROW_EVEN     = "FFFFFF"
    ACCENT       = "048A81"

    if EXCEL_FILE.exists():
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "קשרים בריבוע - מאגר"
        headers = ["תאריך", "מילים שנמצאו", "מילים עבריות", "קטגוריות (JSON)"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font      = Font(bold=True, color="FFFFFF", size=12, name="Arial")
            cell.fill      = PatternFill("solid", fgColor=HEADER_COLOR)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = Border(bottom=Side(style="medium", color=ACCENT), right=Side(style="thin", color="CCCCCC"))
        ws.row_dimensions[1].height = 30
        for col, w in zip(["A","B","C","D"], [15, 40, 50, 60]):
            ws.column_dimensions[col].width = w

    row = ws.max_row + 1
    fill_color = ROW_ODD if row % 2 == 0 else ROW_EVEN
    values = [
        puzzle["date"],
        ", ".join(puzzle.get("all_words", [])[:30]),
        ", ".join(puzzle.get("hebrew_words_found", [])[:40]),
        json.dumps(puzzle.get("categories", []), ensure_ascii=False)[:500],
    ]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.fill      = PatternFill("solid", fgColor=fill_color)
        cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
        cell.font      = Font(name="Arial", size=10)
        cell.border    = Border(bottom=Side(style="thin", color="DDDDDD"), right=Side(style="thin", color="DDDDDD"))
    ws.row_dimensions[row].height = 60
    wb.save(EXCEL_FILE)
    print(f"✅ Excel נשמר")


def log(message: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


async def main():
    log("=" * 50)
    log(f"מתחיל סריקה: {date.today().isoformat()}")
    try:
        raw    = await fetch_puzzle_data()
        puzzle = parse_puzzle(raw)
        log(f"מילים עבריות: {len(puzzle.get('hebrew_words_found', []))}")
        save_to_json(puzzle)
        save_to_excel(puzzle)
        log("✅ הסריקה הושלמה!")
    except Exception as e:
        log(f"❌ שגיאה: {e}")
        raise


if __name__ == "__main__":
    asyncio.run(main())
