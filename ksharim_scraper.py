"""
קשרים בריבוע - Scraper חכם עם Groq AI
לולאת ניחוש אמיתית: הבוט לוחץ, בודק אם נכון, ומתקן אם טעה
"""

import asyncio
import json
import re
import os
import time
import urllib.request
from datetime import date, datetime
from pathlib import Path

from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# הגדרות
# ============================================================
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "הכנסי את ה-KEY שלך כאן")
GROQ_MODEL   = "llama-3.3-70b-versatile"
URL          = "https://ksharim-baribua.com/"

DATA_DIR   = Path("data")
DATA_DIR.mkdir(exist_ok=True)
JSON_FILE  = DATA_DIR / "puzzles_database.json"
EXCEL_FILE = DATA_DIR / "puzzles_database.xlsx"
LOG_FILE   = DATA_DIR / "scrape_log.txt"


# ============================================================
# Groq AI
# ============================================================
def ask_groq(prompt: str) -> str:
    """שולח prompt ל-Groq ומחזיר את התשובה כטקסט."""
    body = json.dumps({
        "model": GROQ_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "max_tokens": 1000,
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.groq.com/openai/v1/chat/completions",
        data=body,
        headers={
            "Content-Type":  "application/json",
            "Authorization": f"Bearer {GROQ_API_KEY}",
        },
        method="POST"
    )

    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode("utf-8"))

    return result["choices"][0]["message"]["content"].strip()


def guess_first_group(remaining_words: list[str], failed_groups: list[list[str]]) -> dict:
    """
    מבקש מ-Groq לנחש את הקבוצה הבטוחה ביותר מתוך המילים שנשארו.
    אם היו קבוצות כושלות - מעביר אותן כדי שלא ינחש שוב.
    """
    failed_str = ""
    if failed_groups:
        failed_str = "\n\nקבוצות שכבר ניסיתי ונכשלו (אל תנחש אותן שוב):\n"
        for f in failed_groups:
            failed_str += f"- {', '.join(f)}\n"

    prompt = f"""אתה משחק "קשרים בריבוע" - משחק עברי שבו צריך למצוא קבוצות של 4 מילים שיש ביניהן קשר.

המילים שנשארו: {', '.join(remaining_words)}
{failed_str}
בחר את הקבוצה שאתה הכי בטוח בה - 4 מילים עם קטגוריה ברורה.

ענה אך ורק ב-JSON תקין, ללא שום טקסט נוסף:
{{"category": "שם הקטגוריה", "words": ["מילה1", "מילה2", "מילה3", "מילה4"]}}"""

    text = ask_groq(prompt)
    text = re.sub(r"```json|```", "", text).strip()
    return json.loads(text)


# ============================================================
# שליפת מילים מהאתר
# ============================================================
async def get_words(page) -> list[str]:
    """שולף את 16 המילים מהדף."""
    await page.wait_for_timeout(3000)

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
        raw     = await page.inner_text("body")
        hebrew  = re.findall(r'[\u05d0-\u05ea]{2,}', raw)
        buttons = list(dict.fromkeys(hebrew))

    words = [w for w in buttons if 1 < len(w) < 15][:16]
    log(f"נמצאו {len(words)} מילים: {', '.join(words)}")
    return words


# ============================================================
# לחיצה על מילים באתר
# ============================================================
async def click_words(page, words: list[str]) -> int:
    """לוחץ על 4 מילים. מחזיר כמה הצליח ללחוץ."""
    clicked = 0
    for word in words:
        try:
            btn = page.get_by_role("button", name=re.compile(f"^{re.escape(word)}$"))
            if await btn.count() > 0:
                await btn.first.click()
                clicked += 1
                await page.wait_for_timeout(400)
                continue
            el = page.locator(f"text={word}").first
            if await el.count() > 0:
                await el.click()
                clicked += 1
                await page.wait_for_timeout(400)
        except Exception:
            pass
    return clicked


async def deselect_all(page):
    """מבטל את כל הבחירות (לחיצה שנייה על כל מילה נבחרת)."""
    try:
        # נסה כפתור "נקה" אם קיים
        clear = page.get_by_role("button", name=re.compile("נקה|ביטול|clear|deselect", re.IGNORECASE))
        if await clear.count() > 0:
            await clear.first.click()
            await page.wait_for_timeout(500)
            return
    except Exception:
        pass

    # אחרת לחץ שוב על כל מילה נבחרת כדי לבטל
    try:
        selected = await page.evaluate("""
            () => {
                const sel = [];
                document.querySelectorAll('[class*="selected"], [class*="active"], [aria-pressed="true"]').forEach(el => {
                    const txt = el.innerText?.trim();
                    if (txt) sel.push(txt);
                });
                return sel;
            }
        """)
        for word in selected:
            try:
                el = page.locator(f"text={word}").first
                if await el.count() > 0:
                    await el.click()
                    await page.wait_for_timeout(300)
            except Exception:
                pass
    except Exception:
        pass


async def check_and_detect_result(page) -> str:
    """
    לוחץ על כפתור 'בדוק' ומזהה אם הניחוש נכון או לא.
    מחזיר: 'correct' / 'wrong' / 'unknown'
    """
    # לחץ בדוק
    try:
        check = page.get_by_role("button", name=re.compile("בדוק|אשר|שלח|submit|confirm", re.IGNORECASE))
        if await check.count() > 0:
            await check.first.click()
            await page.wait_for_timeout(2000)
    except Exception:
        pass

    # בדוק אם יש הודעת הצלחה/כישלון
    result_text = await page.evaluate("""
        () => {
            const body = document.body.innerText.toLowerCase();
            // חפש סימנים להצלחה
            if (body.includes('כל הכבוד') || body.includes('נכון') ||
                body.includes('correct') || body.includes('success') ||
                body.includes('ניצחת') || body.includes('מעולה')) {
                return 'correct';
            }
            // חפש סימנים לכישלון
            if (body.includes('טעות') || body.includes('נסה שוב') ||
                body.includes('wrong') || body.includes('incorrect') ||
                body.includes('לא נכון') || body.includes('חבל')) {
                return 'wrong';
            }
            return 'unknown';
        }
    """)

    # בדוק גם לפי כמה קבוצות נפתרו (אם נפתרה אחת נוספת = נכון)
    solved_count = await page.evaluate("""
        () => {
            let count = 0;
            const selectors = ['[class*="solved"]','[class*="complete"]','[class*="revealed"]','[class*="done"]'];
            for (const sel of selectors) {
                count += document.querySelectorAll(sel).length;
            }
            return count;
        }
    """)

    log(f"  תוצאה: {result_text}, קבוצות שנפתרו: {solved_count}")
    return result_text


# ============================================================
# לולאת הפתרון הראשית
# ============================================================
async def solve_puzzle(page, words_16: list[str]) -> list[dict]:
    """
    לולאה שמנחשת קבוצה אחת בכל פעם:
    - אם נכון: שומרת ועוברת לקבוצה הבאה
    - אם טועה: שולחת ל-Groq שוב עם המידע על הכישלון
    מקסימום 12 ניסיונות סה"כ (3 ניסיונות לכל קבוצה * 4 קבוצות)
    """
    solved_groups  = []   # קבוצות שאושרו
    remaining      = list(words_16)  # מילים שעוד לא נפתרו
    failed_groups  = []   # קבוצות שניסינו ונכשלו
    max_attempts   = 12
    attempt        = 0

    while len(solved_groups) < 4 and attempt < max_attempts and len(remaining) >= 4:
        attempt += 1
        log(f"\n--- ניסיון {attempt} | נפתרו: {len(solved_groups)}/4 | נשארו: {len(remaining)} מילים ---")

        # נחש קבוצה
        try:
            group = guess_first_group(remaining, failed_groups)
            cat   = group.get("category", "")
            words = group.get("words", [])
            log(f"Groq מנחש: {cat} → {words}")
        except Exception as e:
            log(f"שגיאת Groq: {e}")
            break

        # וודא שכל המילים קיימות ברשימה
        words = [w for w in words if w in remaining]
        if len(words) != 4:
            log(f"  ⚠️ לא כל המילים קיימות, מדלג")
            failed_groups.append(group.get("words", []))
            continue

        # לחץ על המילים
        clicked = await click_words(page, words)
        log(f"  לחצתי על {clicked}/4 מילים")

        # בדוק תוצאה
        result = await check_and_detect_result(page)

        if result == "correct":
            log(f"  ✅ נכון! קבוצה: {cat}")
            solved_groups.append({"category": cat, "words": words, "correct": True})
            remaining = [w for w in remaining if w not in words]
            failed_groups = []  # אפס כישלונות לקבוצה הבאה
        else:
            log(f"  ❌ טעות, מנסה שוב עם מידע מעודכן")
            failed_groups.append(words)
            await deselect_all(page)

    # אם נשארו מילים ויש 3 קבוצות - הקבוצה האחרונה ברורה
    if len(solved_groups) == 3 and len(remaining) == 4:
        log("קבוצה אחרונה - מה שנשאר!")
        solved_groups.append({"category": "קבוצה אחרונה", "words": remaining, "correct": True})

    return solved_groups


# ============================================================
# סריקה ראשית
# ============================================================
async def fetch_and_solve() -> dict:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(locale="he-IL")
        page    = await context.new_page()

        await page.goto(URL, wait_until="networkidle", timeout=30000)

        words = await get_words(page)
        if len(words) < 8:
            return {"error": "לא נמצאו מספיק מילים", "words": words, "solved_groups": []}

        log("מתחיל לפתור עם Groq AI...")
        solved = await solve_puzzle(page, words)

        await browser.close()

        return {
            "words_found":   words,
            "solved_groups": solved,
            "total_solved":  len(solved),
        }


# ============================================================
# שמירה ל-JSON
# ============================================================
def save_to_json(puzzle: dict):
    if JSON_FILE.exists():
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            db = json.load(f)
    else:
        db = {"puzzles": [], "total": 0}

    today    = date.today().isoformat()
    existing = [i for i, p in enumerate(db["puzzles"]) if p.get("date") == today]
    if existing:
        db["puzzles"][existing[0]] = puzzle
    else:
        db["puzzles"].append(puzzle)

    db["total"]        = len(db["puzzles"])
    db["last_updated"] = datetime.now().isoformat()

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)
    log("✅ JSON נשמר")


# ============================================================
# שמירה ל-Excel
# ============================================================
def save_to_excel(puzzle: dict):
    HEADER_BG = "2E4057"
    COLORS    = {0: "AED6F1", 1: "ABEBC6", 2: "FAD7A0", 3: "F1948A"}

    if EXCEL_FILE.exists():
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "קשרים בריבוע"
        headers  = ["תאריך", "קטגוריה", "מילה 1", "מילה 2", "מילה 3", "מילה 4", "נכון?"]
        for col, h in enumerate(headers, 1):
            cell           = ws.cell(row=1, column=col, value=h)
            cell.font      = Font(bold=True, color="FFFFFF", size=12, name="Arial")
            cell.fill      = PatternFill("solid", fgColor=HEADER_BG)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 30
        for c in ["C","D","E","F"]:
            ws.column_dimensions[c].width = 14
        ws.column_dimensions["G"].width = 10
        ws.row_dimensions[1].height = 28

    today  = puzzle.get("date", date.today().isoformat())
    groups = puzzle.get("solved_groups", [])

    for i, group in enumerate(groups):
        row   = ws.max_row + 1
        color = COLORS.get(i, "FFFFFF")
        words = group.get("words", [])
        vals  = [
            today,
            group.get("category", ""),
            *words[:4],
            "✅" if group.get("correct") else "❓"
        ]
        for col, val in enumerate(vals, 1):
            cell           = ws.cell(row=row, column=col, value=val)
            cell.fill      = PatternFill("solid", fgColor=color)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font      = Font(name="Arial", size=11)
            cell.border    = Border(
                bottom=Side(style="thin", color="CCCCCC"),
                right=Side(style="thin",  color="CCCCCC")
            )
        ws.row_dimensions[row].height = 24

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
    log(f"סריקה חכמה עם Groq: {date.today().isoformat()}")

    result             = await fetch_and_solve()
    result["date"]     = date.today().isoformat()
    result["fetched_at"] = datetime.now().isoformat()

    save_to_json(result)
    save_to_excel(result)

    log("=" * 50)
    log("📋 תוצאות סופיות:")
    for g in result.get("solved_groups", []):
        status = "✅" if g.get("correct") else "❓"
        log(f"  {status} {g['category']}: {', '.join(g.get('words', []))}")
    log(f"סה\"כ נפתרו: {result.get('total_solved', 0)}/4")
    log("✅ הושלם!")


if __name__ == "__main__":
    asyncio.run(main())
