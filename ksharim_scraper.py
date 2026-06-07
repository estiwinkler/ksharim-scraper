"""
קשרים בריבוע - Scraper חכם עם Gemini AI
הבוט רואה את המילים, מנחש קבוצות בעזרת AI, ושומר את הפתרונות
"""

import asyncio
import json
import re
import os
import urllib.request
from datetime import date, datetime
from pathlib import Path

from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# הגדרות - שני את ה-KEY בלבד!
# ============================================================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "הכנסי את ה-KEY שלך כאן")
URL = "https://ksharim-baribua.com/"

DATA_DIR   = Path("data")
DATA_DIR.mkdir(exist_ok=True)
JSON_FILE  = DATA_DIR / "puzzles_database.json"
EXCEL_FILE = DATA_DIR / "puzzles_database.xlsx"
LOG_FILE   = DATA_DIR / "scrape_log.txt"


# ============================================================
# Gemini AI - ניחוש קבוצות
# ============================================================
def ask_gemini(words: list[str]) -> list[dict]:
    """שולח את 16 המילים ל-Gemini ומקבל 4 קבוצות עם קטגוריות."""
    words_str = ", ".join(words)
    prompt = f"""אתה משחק "קשרים בריבוע" - משחק עברי שבו צריך למצוא 4 קבוצות של 4 מילים שיש ביניהן קשר.

הנה 16 המילים: {words_str}

חלק אותן ל-4 קבוצות של 4 מילים. לכל קבוצה תן שם קטגוריה קצר שמסביר את הקשר.

ענה אך ורק ב-JSON תקין בפורמט הזה בדיוק, ללא שום טקסט נוסף:
[
  {{"category": "שם הקטגוריה", "words": ["מילה1", "מילה2", "מילה3", "מילה4"]}},
  {{"category": "שם הקטגוריה", "words": ["מילה1", "מילה2", "מילה3", "מילה4"]}},
  {{"category": "שם הקטגוריה", "words": ["מילה1", "מילה2", "מילה3", "מילה4"]}},
  {{"category": "שם הקטגוריה", "words": ["מילה1", "מילה2", "מילה3", "מילה4"]}}
]"""

    body = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1}
    }).encode("utf-8")

    # נסה מספר מודלים בסדר עדיפות
    models = [
        "gemini-1.5-flash",
        "gemini-1.5-flash-8b",
        "gemini-1.0-pro",
    ]

    last_error = None
    for model in models:
        try:
            log(f"מנסה מודל: {model}")
            req = urllib.request.Request(
                f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GEMINI_API_KEY}",
                data=body,
                headers={"Content-Type": "application/json"},
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=30) as resp:
                result = json.loads(resp.read().decode("utf-8"))
            text = result["candidates"][0]["content"]["parts"][0]["text"]
            text = re.sub(r"```json|```", "", text).strip()
            parsed = json.loads(text)
            log(f"✅ הצליח עם מודל: {model}")
            return parsed
        except Exception as e:
            log(f"מודל {model} נכשל: {e}")
            last_error = e
            import time
            time.sleep(2)  # המתן 2 שניות לפני הניסיון הבא

    raise Exception(f"כל המודלים נכשלו. שגיאה אחרונה: {last_error}")


# ============================================================
# Playwright - שליפת מילים ולחיצה
# ============================================================
async def get_words_and_solve(page) -> dict:
    """שולף את 16 המילים, שולח ל-Gemini, ולוחץ לפי הניחוש."""

    # --- שלב 1: שלוף את 16 המילים ---
    log("מחפש את 16 המילים...")
    await page.wait_for_timeout(3000)

    # נסה למצוא כפתורי מילים
    buttons = await page.evaluate("""
        () => {
            const btns = [];
            document.querySelectorAll('button, [role="button"], [class*="word"], [class*="Word"], [class*="tile"], [class*="card"]').forEach(el => {
                const txt = el.innerText?.trim();
                if (txt && /[\u05d0-\u05ea]{2,}/.test(txt) && txt.length < 20) {
                    btns.push(txt);
                }
            });
            return [...new Set(btns)];
        }
    """)

    if len(buttons) < 12:
        # נסה גישה שנייה - כל הטקסט העברי
        raw = await page.inner_text("body")
        hebrew = re.findall(r'[\u05d0-\u05ea]{2,}', raw)
        buttons = list(dict.fromkeys(hebrew))

    # קח רק 16 הראשונים אם יש יותר
    words_16 = [w for w in buttons if 1 < len(w) < 15][:16]
    log(f"נמצאו {len(words_16)} מילים: {', '.join(words_16)}")

    if len(words_16) < 8:
        return {"error": "לא נמצאו מספיק מילים", "words": words_16, "groups": []}

    # --- שלב 2: שלח ל-Gemini ---
    log("שולח ל-Gemini AI לניחוש קבוצות...")
    try:
        ai_groups = ask_gemini(words_16)
        log(f"Gemini הציע {len(ai_groups)} קבוצות")
    except Exception as e:
        log(f"שגיאת Gemini: {e}")
        ai_groups = []

    # --- שלב 3: נסה ללחוץ על כל קבוצה ---
    confirmed_groups = []
    for group in ai_groups:
        cat   = group.get("category", "")
        words = group.get("words", [])
        log(f"מנסה קבוצה: {cat} → {words}")

        # לחץ על 4 המילים
        clicked = 0
        for word in words:
            try:
                # נסה ללחוץ על הכפתור עם הטקסט הזה
                btn = page.get_by_role("button", name=re.compile(word, re.IGNORECASE))
                if await btn.count() > 0:
                    await btn.first.click()
                    clicked += 1
                    await page.wait_for_timeout(300)
                else:
                    # נסה בדרך אחרת
                    el = page.locator(f"text={word}").first
                    if await el.count() > 0:
                        await el.click()
                        clicked += 1
                        await page.wait_for_timeout(300)
            except Exception:
                pass

        log(f"  לחצתי על {clicked}/4 מילים")

        # לחץ "בדוק" אם קיים
        try:
            check_btn = page.get_by_role("button", name=re.compile("בדוק|אשר|שלח|confirm|submit", re.IGNORECASE))
            if await check_btn.count() > 0:
                await check_btn.first.click()
                await page.wait_for_timeout(1500)
        except Exception:
            pass

        confirmed_groups.append({
            "category": cat,
            "words":    words,
            "clicked":  clicked
        })

    # --- שלב 4: שלוף קבוצות שנחשפו בדף ---
    revealed = await page.evaluate("""
        () => {
            const groups = [];
            const selectors = ['[class*="group"]','[class*="Group"]','[class*="category"]','[class*="solved"]','[class*="complete"]'];
            for (const sel of selectors) {
                document.querySelectorAll(sel).forEach(el => {
                    const txt = el.innerText?.trim();
                    if (txt && txt.length > 3) groups.push(txt);
                });
            }
            return groups;
        }
    """)

    return {
        "words_found":       words_16,
        "ai_suggested":      ai_groups,
        "confirmed_groups":  confirmed_groups,
        "revealed_on_page":  revealed,
    }


# ============================================================
# סריקה ראשית
# ============================================================
async def fetch_and_solve() -> dict:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(locale="he-IL")
        page    = await context.new_page()

        await page.goto(URL, wait_until="networkidle", timeout=30000)
        result = await get_words_and_solve(page)
        await browser.close()
        return result


# ============================================================
# שמירה ל-JSON
# ============================================================
def save_to_json(puzzle: dict):
    if JSON_FILE.exists():
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            db = json.load(f)
    else:
        db = {"puzzles": [], "total": 0}

    today = date.today().isoformat()
    existing = [i for i, p in enumerate(db["puzzles"]) if p.get("date") == today]
    if existing:
        db["puzzles"][existing[0]] = puzzle
    else:
        db["puzzles"].append(puzzle)

    db["total"] = len(db["puzzles"])
    db["last_updated"] = datetime.now().isoformat()

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)
    log("✅ JSON נשמר")


# ============================================================
# שמירה ל-Excel
# ============================================================
def save_to_excel(puzzle: dict):
    HEADER_BG = "2E4057"
    COLORS = {
        0: "AED6F1",  # כחול
        1: "ABEBC6",  # ירוק
        2: "FAD7A0",  # כתום
        3: "F1948A",  # אדום
    }

    if EXCEL_FILE.exists():
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws  = wb.active
        ws.title = "קשרים בריבוע"
        headers = ["תאריך", "קטגוריה", "מילה 1", "מילה 2", "מילה 3", "מילה 4"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font      = Font(bold=True, color="FFFFFF", size=12, name="Arial")
            cell.fill      = PatternFill("solid", fgColor=HEADER_BG)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 30
        for c in ["C","D","E","F"]:
            ws.column_dimensions[c].width = 16
        ws.row_dimensions[1].height = 28

    today  = puzzle.get("date", date.today().isoformat())
    groups = puzzle.get("ai_suggested", [])

    for i, group in enumerate(groups):
        row   = ws.max_row + 1
        color = COLORS.get(i, "FFFFFF")
        words = group.get("words", [])
        vals  = [today, group.get("category",""), *words[:4]]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill      = PatternFill("solid", fgColor=color)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font      = Font(name="Arial", size=11)
            cell.border    = Border(
                bottom=Side(style="thin", color="CCCCCC"),
                right=Side(style="thin",  color="CCCCCC")
            )
        ws.row_dimensions[row].height = 24

    # שורה ריקה בין ימים
    ws.append([""])
    wb.save(EXCEL_FILE)
    log("✅ Excel נשמר")


# ============================================================
# לוג
# ============================================================
def log(message: str):
    ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ============================================================
# Main
# ============================================================
async def main():
    log("=" * 50)
    log(f"סריקה חכמה: {date.today().isoformat()}")

    result = await fetch_and_solve()
    result["date"] = date.today().isoformat()
    result["fetched_at"] = datetime.now().isoformat()

    save_to_json(result)
    save_to_excel(result)

    log("=" * 50)
    log("📋 תוצאות:")
    for g in result.get("ai_suggested", []):
        log(f"  🔹 {g['category']}: {', '.join(g.get('words',[]))}")
    log("✅ הושלם!")


if __name__ == "__main__":
    asyncio.run(main())
