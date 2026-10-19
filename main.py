
import os
import json
import re
import time
from pathlib import Path
from urllib.parse import urlparse
from datetime import datetime

import requests
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash").strip()

OUT = Path("output")
OUT.mkdir(exist_ok=True)

TARGET_COUNT = 300
BATCH_SIZE = 10
TOTAL_BATCHES = TARGET_COUNT // BATCH_SIZE

PROGRESS_FILE = OUT / "manager_test_progress.json"
INVENTORY_FILE = OUT / "manual_discovery_inventory.json"
WORKBOOK_FILE = OUT / "Manager_Role_300_Test_Cases.xlsx"

SKIP_WORDS = (
    "logout", "log-out", "signout", "sign-out",
    "delete", "remove", "archive", "terminate"
)

HEADERS = [
    "Test Case ID", "Module", "Parent Feature", "Child/Sub-feature",
    "Related Features", "Test Scenario", "Test Steps", "Expected Result",
    "Actual Result", "Preconditions", "Test Type", "Severity", "Priority",
    "Status", "Bug ID", "Comments", "Source URL"
]


def read_json(path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return default


def write_json(path, data):
    temp = path.with_suffix(".tmp")
    temp.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    temp.replace(path)


def snapshot(page):
    """Capture visible page structure without submitting business actions."""
    return page.evaluate("""() => {
      const clean = s => (s || '').replace(/\\s+/g, ' ').trim();
      const visible = e => {
        const r = e.getBoundingClientRect();
        const s = getComputedStyle(e);
        return r.width > 0 && r.height > 0 &&
               s.display !== 'none' && s.visibility !== 'hidden';
      };
      const all = (sel, fn, n=100) =>
        Array.from(document.querySelectorAll(sel))
          .filter(visible).slice(0,n).map(fn);

      return {
        url: location.href,
        title: document.title,
        headings: all('h1,h2,h3,[role=heading]',
          e => clean(e.innerText), 80),
        links: all('a[href]', e => ({
          text: clean(e.innerText || e.getAttribute('aria-label')),
          href: e.href
        }), 180),
        buttons: all('button,[role=button]', e => ({
          text: clean(e.innerText || e.getAttribute('aria-label')),
          disabled: !!e.disabled,
          expanded: e.getAttribute('aria-expanded')
        }), 150),
        tabs: all('[role=tab],nav a,aside a,[role=menuitem]',
          e => clean(e.innerText || e.getAttribute('aria-label')), 150),
        fields: all('input,textarea,select,[contenteditable=true]',
          e => ({
            tag: e.tagName.toLowerCase(),
            type: e.type || '',
            name: e.name || '',
            label: clean(
              (e.labels &&
               Array.from(e.labels).map(x => x.innerText).join(' '))
              || e.getAttribute('aria-label')
              || e.getAttribute('title')
            ),
            placeholder: e.placeholder || '',
            required: !!e.required,
            disabled: !!e.disabled,
            value: e.tagName === 'SELECT'
              ? Array.from(e.options).map(o => clean(o.text)).slice(0,40)
              : undefined
          }), 150),
        tables: all('table,[role=grid]', e => clean(e.innerText).slice(0,2000), 15),
        text: clean(document.body.innerText).slice(0,12000)
      };
    }""")


def capture_current_page(page, feature_name, captures):
    """Record the current page and save discovery progress immediately."""
    try:
        page.wait_for_load_state("domcontentloaded", timeout=5000)
    except Exception:
        pass

    data = snapshot(page)
    data["url"] = page.url
    data["feature_name"] = feature_name.strip()

    signature = (
        data["feature_name"].lower(),
        data["url"],
        tuple(data.get("headings", [])),
        tuple(data.get("tabs", [])),
        tuple(
            (f.get("label"), f.get("name"), f.get("type"), f.get("placeholder"))
            for f in data.get("fields", [])
        ),
        data.get("text", "")[:1500]
    )

    for existing in captures:
        existing_signature = (
            existing.get("feature_name", "").lower(),
            existing.get("url", ""),
            tuple(existing.get("headings", [])),
            tuple(existing.get("tabs", [])),
            tuple(
                (f.get("label"), f.get("name"), f.get("type"), f.get("placeholder"))
                for f in existing.get("fields", [])
            ),
            existing.get("text", "")[:1500]
        )
        if signature == existing_signature:
            print("This page state is already captured; not adding a duplicate.")
            return False

    captures.append(data)
    write_json(INVENTORY_FILE, captures)

    print(f"Captured: {data.get('title', 'Page')} | {data['url']}")
    print(f"Headings: {len(data.get('headings', []))}")
    print(f"Fields: {len(data.get('fields', []))}")
    print(f"Buttons: {len(data.get('buttons', []))}")
    print(f"Saved discovery inventory: {INVENTORY_FILE}")
    return True


def call_gemini(prompt):
    if not API_KEY:
        raise RuntimeError("GEMINI_API_KEY is missing from your .env file.")

    url = (
        "https://generativelanguage.googleapis.com/v1beta/"
        f"models/{MODEL}:generateContent?key={API_KEY}"
    )

    last_error = None

    for attempt in range(5):
        try:
            response = requests.post(
                url,
                json={
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {
                        "temperature": 0.2,
                        "responseMimeType": "application/json"
                    }
                },
                timeout=180
            )

            if response.status_code == 200:
                payload = response.json()
                raw = payload["candidates"][0]["content"]["parts"][0]["text"]
                raw = re.sub(
                    r"^\s*```(?:json)?\s*|\s*```\s*$",
                    "",
                    raw.strip(),
                    flags=re.I
                )
                return json.loads(raw)

            last_error = (
                f"Gemini HTTP {response.status_code}: "
                f"{response.text[:700]}"
            )

            if response.status_code not in (429, 500, 502, 503, 504):
                raise RuntimeError(last_error)

            print(f"{last_error} — attempt {attempt + 1}/5")

        except (requests.RequestException, json.JSONDecodeError, KeyError) as exc:
            last_error = str(exc)
            print(f"API/response error — attempt {attempt + 1}/5: {last_error}")

        if attempt < 4:
            time.sleep(min(5 * (2 ** attempt), 60))

    raise RuntimeError(
        f"Could not generate the batch after 5 attempts. {last_error}"
    )


def normalize_case(case, feature_name, source_url):
    """Keep required fields present and normalize common output mistakes."""
    return {
        "Module": str(case.get("Module") or feature_name),
        "Parent Feature": str(case.get("Parent Feature") or feature_name),
        "Child/Sub-feature": str(case.get("Child/Sub-feature") or "General"),
        "Related Features": str(case.get("Related Features") or ""),
        "Test Scenario": str(case.get("Test Scenario") or "").strip(),
        "Test Steps": str(case.get("Test Steps") or "").strip(),
        "Expected Result": str(case.get("Expected Result") or "").strip(),
        "Actual Result": "Not Executed",
        "Preconditions": str(case.get("Preconditions") or "Manager role is logged in."),
        "Test Type": str(case.get("Test Type") or "Functional"),
        "Severity": str(case.get("Severity") or "Medium"),
        "Priority": str(case.get("Priority") or "P2"),
        "Source URL": str(case.get("Source URL") or source_url),
        "Comments": str(case.get("Comments") or
                        "AI-generated; review assumptions before execution.")
    }


def case_signature(case):
    return (
        case.get("Module", "").strip().lower(),
        case.get("Parent Feature", "").strip().lower(),
        case.get("Child/Sub-feature", "").strip().lower(),
        case.get("Test Scenario", "").strip().lower()
    )


def generate_batch(feature_name, captured_pages, batch_number, previous_cases):
    """Generate one batch of 30 cases for the manually selected feature."""
    if not captured_pages:
        raise RuntimeError("No pages captured for this batch.")

    # Keep the current batch's inventory focused and within a practical prompt size.
    inventory = []
    for page in captured_pages:
        inventory.append({
            "feature_name": page.get("feature_name", feature_name),
            "url": page.get("url", ""),
            "title": page.get("title", ""),
            "headings": page.get("headings", [])[:60],
            "tabs": page.get("tabs", [])[:100],
            "buttons": page.get("buttons", [])[:100],
            "fields": page.get("fields", [])[:100],
            "tables": page.get("tables", [])[:12],
            "text": page.get("text", "")[:5000]
        })

    # Send only the existing scenario names to discourage repetition.
    previous_scenarios = [
        c.get("Test Scenario", "")
        for c in previous_cases
        if c.get("Test Scenario")
    ][-300:]

    prompt = f"""
You are a senior software QA engineer preparing detailed manual test cases
for a workforce/HR web application, using the logged-in Manager role.

TASK:
Generate exactly {BATCH_SIZE} distinct, detailed manual test cases for:
Feature / parent feature: {feature_name}
Batch number: {batch_number} of {TOTAL_BATCHES}

Use ONLY the observed inventory below. Do not invent pages, buttons, fields,
permissions, business rules, or integrations that are not evidenced.

COVERAGE:
- Parent feature and all observed child/sub-features.
- Normal/positive workflows, required-field validation, invalid input,
  boundary values, field formats, dropdowns, filters, search, sorting,
  pagination, tables, empty states, error handling, navigation and UI state.
- Include role/access checks only when the available UI or inventory supports
  them; otherwise do not assume hidden permissions.
- Include accessibility/keyboard checks only where practical for the observed UI.
- Cover different scenarios, not just the same action with different wording.
- Avoid all scenarios that duplicate these existing cases:
{json.dumps(previous_scenarios, ensure_ascii=False)}

DETAIL REQUIREMENTS:
- Test Scenario must describe one testable behavior.
- Test Steps must be numbered, specific and executable by a human tester.
- Include setup/navigation, test data where relevant, the action, and what to inspect.
- Expected Result must be observable and specific.
- Do not claim a test passed. Actual Result must be "Not Executed".
- If a business rule cannot be confirmed from the inventory, begin the relevant
  expected result with "Assumption to verify:".
- Use the feature name in Module or Parent Feature.
- Use a meaningful Child/Sub-feature based on the observed UI.
- Priority must be P1, P2, or P3.
- Severity must be High, Medium, or Low.
- Test Type should be Functional, UI, Negative, Boundary, Accessibility, or Integration.
- Source URL must be one of the captured URLs.
- Do not number test case IDs; the program assigns them.

Return only valid JSON in this format:
{{
  "cases": [
    {{
      "Module": "...",
      "Parent Feature": "...",
      "Child/Sub-feature": "...",
      "Related Features": "...",
      "Test Scenario": "...",
      "Test Steps": "1. ...\\n2. ...\\n3. ...",
      "Expected Result": "...",
      "Actual Result": "Not Executed",
      "Preconditions": "...",
      "Test Type": "Functional",
      "Severity": "Medium",
      "Priority": "P2",
      "Source URL": "...",
      "Comments": "..."
    }}
  ]
}}

CAPTURED INVENTORY:
{json.dumps(inventory, ensure_ascii=False)}
"""

    result = call_gemini(prompt)
    raw_cases = result.get("cases", [])

    if not isinstance(raw_cases, list):
        raise RuntimeError("The model response did not contain a cases list.")

    normalized = []
    seen = {case_signature(c) for c in previous_cases}

    for item in raw_cases:
        if not isinstance(item, dict):
            continue
        case = normalize_case(
            item,
            feature_name,
            inventory[0].get("url", "")
        )

        if not case["Test Scenario"] or not case["Test Steps"] or not case["Expected Result"]:
            continue

        sig = case_signature(case)
        if sig in seen:
            continue

        seen.add(sig)
        normalized.append(case)

    if len(normalized) < BATCH_SIZE:
        raise RuntimeError(
            f"Gemini returned only {len(normalized)} valid, non-duplicate cases "
            f"for this batch. Nothing was marked complete. Try again with more "
            f"captured sub-features, or adjust the feature name/inventory."
        )

    return normalized[:BATCH_SIZE]


def load_existing_cases():
    progress = read_json(PROGRESS_FILE, {"completed_batches": [], "cases": []})
    return progress


def save_progress(progress):
    write_json(PROGRESS_FILE, progress)
    write_json(INVENTORY_FILE, progress.get("captures", []))


def save_xlsx(cases):
    """Rewrite the workbook safely after each completed batch."""
    temp_path = WORKBOOK_FILE.with_name("Manager_Role_300_Test_Cases.tmp.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Manager Test Cases"
    ws.append(HEADERS)

    for i, case in enumerate(cases, 1):
        ws.append([
            f"TC-{i:03d}",
            case.get("Module", ""),
            case.get("Parent Feature", ""),
            case.get("Child/Sub-feature", ""),
            case.get("Related Features", ""),
            case.get("Test Scenario", ""),
            case.get("Test Steps", ""),
            case.get("Expected Result", ""),
            case.get("Actual Result", "Not Executed"),
            case.get("Preconditions", ""),
            case.get("Test Type", ""),
            case.get("Severity", "Medium"),
            case.get("Priority", "P2"),
            "Not Run",
            "",
            case.get("Comments", ""),
            case.get("Source URL", "")
        ])

    for cell in ws[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    widths = [14, 22, 24, 28, 28, 42, 65, 60, 20, 38, 22, 14, 12, 14, 14, 42, 48]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32

    # Add one worksheet per feature batch for easier review.
    for batch in sorted(
        {c.get("_batch_number") for c in cases if c.get("_batch_number")}
    ):
        batch_cases = [c for c in cases if c.get("_batch_number") == batch]
        feature = batch_cases[0].get("_batch_feature", f"Batch {batch}")
        sheet_name = f"B{batch:02d}_{feature}"[:31]
        sheet_name = re.sub(r"[\[\]:*?/\\\\]", "_", sheet_name)

        # Ensure unique worksheet names.
        original = sheet_name
        suffix = 1
        while sheet_name in wb.sheetnames:
            sheet_name = f"{original[:27]}_{suffix}"
            suffix += 1

        batch_ws = wb.create_sheet(sheet_name)
        batch_ws.append(HEADERS)
        for i, case in enumerate(batch_cases, 1):
            batch_ws.append([
                f"TC-{(batch - 1) * BATCH_SIZE + i:03d}",
                case.get("Module", ""),
                case.get("Parent Feature", ""),
                case.get("Child/Sub-feature", ""),
                case.get("Related Features", ""),
                case.get("Test Scenario", ""),
                case.get("Test Steps", ""),
                case.get("Expected Result", ""),
                case.get("Actual Result", "Not Executed"),
                case.get("Preconditions", ""),
                case.get("Test Type", ""),
                case.get("Severity", "Medium"),
                case.get("Priority", "P2"),
                "Not Run",
                "",
                case.get("Comments", ""),
                case.get("Source URL", "")
            ])
        for cell in batch_ws[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for i, width in enumerate(widths, 1):
            batch_ws.column_dimensions[get_column_letter(i)].width = width
        for row in batch_ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        batch_ws.freeze_panes = "A2"
        batch_ws.auto_filter.ref = batch_ws.dimensions

    wb.save(temp_path)
    temp_path.replace(WORKBOOK_FILE)
    return WORKBOOK_FILE


def main():
    if not API_KEY:
        raise SystemExit(
            "Missing GEMINI_API_KEY. Add GEMINI_API_KEY=your_key to your .env file."
        )

    progress = load_existing_cases()
    progress.setdefault("completed_batches", [])
    progress.setdefault("cases", [])
    progress.setdefault("captures", [])

    print("\n=== Manual Feature-by-Feature QA Test Case Generator ===")
    print(f"Target: {TARGET_COUNT} cases, {BATCH_SIZE} per batch.")
    print(f"Completed batches: {progress['completed_batches']}")
    print(f"Existing cases: {len(progress['cases'])}")
    print(f"Workbook: {WORKBOOK_FILE.resolve()}")
    print("\nCommands:")
    print("  Enter     Capture the currently visible page")
    print("  GEN       Generate 30 cases for the current feature")
    print("  DONE      Finish without generating the current feature")
    print("  QUIT      Save and exit")
    print("\nIn Chrome, navigate manually to each feature/sub-feature.")
    print("Capture every relevant screen for a feature before entering GEN.")
    print("Do not change features while a batch is being generated.")

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        except Exception as exc:
            raise SystemExit(
                "Could not connect to Chrome. Run start_chrome.bat first. "
                + str(exc)
            )

        if not browser.contexts or not browser.contexts[0].pages:
            raise SystemExit(
                "Open the dedicated Chrome window and log in first."
            )

        page = browser.contexts[0].pages[-1]
        print(f"\nConnected to current page: {page.url}")

        while len(progress["cases"]) < TARGET_COUNT:
            next_batch = next(
                (
                    n for n in range(1, TOTAL_BATCHES + 1)
                    if n not in progress["completed_batches"]
                ),
                None
            )

            if next_batch is None:
                break

            print("\n" + "=" * 65)
            print(
                f"BATCH {next_batch}/{TOTAL_BATCHES} "
                f"— {len(progress['cases'])}/{TARGET_COUNT} cases complete"
            )
            feature_name = input(
                "Enter the parent feature name for this batch "
                "(e.g. Team Leave): "
            ).strip()

            if not feature_name:
                print("Feature name cannot be empty.")
                continue

            captures = []
            print(f"\nFeature: {feature_name}")
            print("Navigate in Chrome and capture its pages/sub-features.")
            print("Press Enter to capture; type GEN when ready.")

            while True:
                command = input(
                    f"\nCapture command [{len(captures)} pages captured]: "
                ).strip()

                if command.upper() == "QUIT":
                    progress["captures"] = captures
                    save_progress(progress)
                    if progress["cases"]:
                        save_xlsx(progress["cases"])
                    print("Progress saved. Exiting.")
                    return

                if command.upper() == "DONE":
                    print("Feature capture ended without generating a batch.")
                    break

                if command.upper() == "GEN":
                    if not captures:
                        print("Capture at least one page before generating.")
                        continue

                    print(f"Generating {BATCH_SIZE} cases for {feature_name}...")
                    try:
                        new_cases = generate_batch(
                            feature_name,
                            captures,
                            next_batch,
                            progress["cases"]
                        )
                    except Exception as exc:
                        print(f"\nBatch failed: {exc}")
                        print("Your previously completed batches remain saved.")
                        retry = input("Retry this batch? (y/n): ").strip().lower()
                        if retry == "y":
                            continue
                        progress["captures"] = captures
                        save_progress(progress)
                        if progress["cases"]:
                            save_xlsx(progress["cases"])
                        return

                    for case in new_cases:
                        case["_batch_number"] = next_batch
                        case["_batch_feature"] = feature_name

                    progress["cases"].extend(new_cases)
                    progress["completed_batches"].append(next_batch)
                    progress["captures"] = captures
                    progress["last_updated"] = datetime.now().isoformat(timespec="seconds")

                    # Save JSON progress and Excel immediately after every batch.
                    save_progress(progress)
                    workbook = save_xlsx(progress["cases"])

                    print(
                        f"\nBATCH SAVED: {len(new_cases)} new cases."
                        f"\nTotal: {len(progress['cases'])}/{TARGET_COUNT}"
                        f"\nExcel: {workbook.resolve()}"
                    )
                    break

                # Blank input captures the current page; any other text also
                # captures the page so accidental commands don't lose discovery.
                capture_current_page(page, feature_name, captures)

            if len(progress["cases"]) >= TARGET_COUNT:
                break

        if progress["cases"]:
            workbook = save_xlsx(progress["cases"])
            print("\nFinished.")
            print(f"Total generated: {len(progress['cases'])}")
            print(f"Excel workbook: {workbook.resolve()}")
            print("Actual Result remains 'Not Executed'; Status remains 'Not Run'.")
        else:
            print("No test cases were generated.")

        print("\nKeep the progress JSON files if you want to resume later.")


if __name__ == "__main__":
    main()
