from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


TEST_HEADERS = [
    "Test Case ID", "Test Scenario", "Preconditions", "Test Steps",
    "Expected Result", "Priority", "Test Type", "Source Evidence",
    "Automation Readiness", "Notes"
]


def _add_case(cases: list[dict[str, str]], scenario: str, pre: str, steps: str,
              expected: str, priority: str, test_type: str, evidence: str,
              readiness: str = "Needs review", notes: str = "") -> None:
    cases.append({
        "Test Case ID": f"TC{len(cases) + 1:03d}",
        "Test Scenario": scenario,
        "Preconditions": pre,
        "Test Steps": steps,
        "Expected Result": expected,
        "Priority": priority,
        "Test Type": test_type,
        "Source Evidence": evidence,
        "Automation Readiness": readiness,
        "Notes": notes,
    })


def build_test_cases(observations: dict[str, Any]) -> list[dict[str, str]]:
    cases: list[dict[str, str]] = []
    final_url = observations.get("final_url", "")
    title = observations.get("title", "")
    status = observations.get("status")

    _add_case(
        cases, "Verify that the website is reachable",
        "Target URL is accessible and the tester is authorized to inspect it.",
        f"1. Navigate to {observations.get('requested_url', '')}.",
        f"Page navigation completes. Observed HTTP status: {status if status is not None else 'not available'}.",
        "High", "Smoke", f"Final URL: {final_url}",
        "Candidate for automation", "A successful navigation does not by itself prove all page features work."
    )
    _add_case(
        cases, "Verify the browser reaches the expected page URL",
        "Website navigation has completed.",
        f"1. Open {observations.get('requested_url', '')}. 2. Read the final browser URL.",
        f"Browser URL matches the intended destination or an approved redirect. Observed: {final_url}.",
        "Medium", "Functional", f"Requested URL: {observations.get('requested_url', '')}; final URL: {final_url}",
        "Candidate for automation"
    )
    if title:
        _add_case(
            cases, "Verify page title is present",
            "Website is loaded.",
            "1. Read the browser page title.",
            f"Page title is not empty and equals the observed title: {title}.",
            "Medium", "UI", f"Observed document title: {title}",
            "Candidate for automation"
        )
    else:
        _add_case(
            cases, "Check whether the page title is defined",
            "Website is loaded.",
            "1. Read the browser page title.",
            "A meaningful, non-empty page title is defined.",
            "Low", "Accessibility / SEO", "No title was returned during inspection.",
            "Candidate for automation", "Generated as a gap check because no title was observed."
        )

    headings = observations.get("headings", [])
    if headings:
        for idx, heading in enumerate(headings[:10], 1):
            label = heading.get("text", "")
            _add_case(
                cases, f"Verify visible heading: {label[:90]}",
                "Website is loaded.",
                f"1. Locate the observed {heading.get('tag', 'heading')} element. 2. Verify it is visible.",
                f"The heading '{label}' is visible on the page.",
                "Low", "UI", f"{heading.get('tag', '')}: {label}",
                "Candidate for automation"
            )

    for idx, link in enumerate(observations.get("links", [])[:30], 1):
        text = link.get("text") or "(no visible link text)"
        href = link.get("href") or "(no href)"
        _add_case(
            cases, f"Verify observed link is visible: {text[:70]}",
            "Website is loaded.",
            f"1. Locate the link with text/accessible name '{text}'. 2. Verify it is visible. 3. Inspect its destination.",
            f"Link is visible and has an appropriate destination. Observed href: {href}.",
            "Low", "UI / Navigation", f"Link text: {text}; href: {href}",
            "Candidate for automation" if link.get("href") else "Needs review",
            "Destination behavior is not tested; this generator does not click links."
        )

    for idx, button in enumerate(observations.get("buttons", [])[:25], 1):
        label = button.get("text") or "(unnamed button)"
        _add_case(
            cases, f"Verify observed button is displayed: {label[:70]}",
            "Website is loaded.",
            f"1. Locate the button '{label}'. 2. Verify visibility and enabled/disabled state.",
            f"Button is displayed. Observed disabled state: {button.get('disabled')}.",
            "Low", "UI", f"Button: {label}; type: {button.get('type')}; disabled: {button.get('disabled')}",
            "Candidate for automation",
            "Button is not clicked because its action may change data or trigger navigation."
        )

    fields = []
    fields.extend([("input", x) for x in observations.get("inputs", [])])
    fields.extend([("textarea", x) for x in observations.get("textareas", [])])
    fields.extend([("select", x) for x in observations.get("selects", [])])
    for field_type, field in fields[:40]:
        label = field.get("label") or field.get("name") or field.get("placeholder") or "(unlabelled field)"
        detail = f"type={field.get('type', field_type)}; required={field.get('required')}; disabled={field.get('disabled')}"
        _add_case(
            cases, f"Verify form field is displayed: {label[:70]}",
            "Website is loaded.",
            f"1. Locate the observed {field_type} field '{label}'. 2. Verify it is visible. 3. Check its required/disabled attributes.",
            f"Field is visible and its attributes match the observed state ({detail}).",
            "Medium", "UI / Form", f"{field_type}: {label}; {detail}",
            "Candidate for automation",
            "No values are entered and no form is submitted."
        )

    if observations.get("forms"):
        for form in observations["forms"][:15]:
            name = form.get("name") or form.get("id") or "(unnamed form)"
            _add_case(
                cases, f"Verify observed form structure: {name}",
                "Website is loaded.",
                f"1. Locate the form '{name}'. 2. Verify it is present. 3. Count its visible controls.",
                f"Form is present and contains the observed {form.get('field_count')} form control(s).",
                "Medium", "UI / Form",
                f"Form name={form.get('name')}; id={form.get('id')}; method={form.get('method')}; fields={form.get('field_count')}",
                "Candidate for automation",
                "Form submission is intentionally not attempted."
            )
    else:
        _add_case(
            cases, "Check whether a form is present",
            "Website is loaded.",
            "1. Inspect the page for form elements.",
            "A form is present if the page is intended to collect user input; otherwise document that no form is expected.",
            "Low", "UI / Form", "No visible form element was detected.",
            "Needs review", "This is a review prompt, not a claim that the website requires a form."
        )

    if not cases:
        _add_case(cases, "Review page content manually", "Website is loaded.",
                  "1. Review the rendered page.", "Relevant page content is displayed.",
                  "Low", "Exploratory", "No supported UI elements were detected.",
                  "Needs review")
    return cases


def _style_sheet(ws, freeze: str = "A2") -> None:
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = max((len(str(ws.cell(r, col_idx).value or "")) for r in range(1, min(ws.max_row, 80) + 1)), default=12)
        ws.column_dimensions[letter].width = min(max(max_len + 2, 14), 48)


def write_workbook(observations: dict[str, Any], output_path: str | Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    cases = build_test_cases(observations)

    wb = Workbook()
    ws = wb.active
    ws.title = "Test Cases"
    ws.append(TEST_HEADERS)
    for case in cases:
        ws.append([case[h] for h in TEST_HEADERS])
    _style_sheet(ws)
    if ws.max_row > 1:
        tab = Table(displayName="GeneratedTestCases", ref=ws.dimensions)
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True,
                                            showColumnStripes=False)
        ws.add_table(tab)

    obs_ws = wb.create_sheet("Page Observations")
    obs_ws.append(["Element Type", "Observed Text / Label", "Details"])
    for heading in observations.get("headings", []):
        obs_ws.append(["Heading", heading.get("text"), heading.get("tag")])
    for link in observations.get("links", []):
        obs_ws.append(["Link", link.get("text"), f"href={link.get('href')}; target={link.get('target')}"])
    for button in observations.get("buttons", []):
        obs_ws.append(["Button", button.get("text"), f"type={button.get('type')}; disabled={button.get('disabled')}"])
    for kind in ("inputs", "textareas", "selects"):
        for field in observations.get(kind, []):
            obs_ws.append([kind[:-1].title(), field.get("label") or field.get("name") or field.get("placeholder"),
                           json_safe(field)])
    for form in observations.get("forms", []):
        obs_ws.append(["Form", form.get("name") or form.get("id"), json_safe(form)])
    _style_sheet(obs_ws)

    info = wb.create_sheet("Run Info")
    info.append(["Property", "Value"])
    rows = [
        ("Requested URL", observations.get("requested_url")),
        ("Final URL", observations.get("final_url")),
        ("Page title", observations.get("title")),
        ("HTTP status", observations.get("status")),
        ("Inspection timestamp (UTC)", datetime.now(timezone.utc).isoformat(timespec="seconds")),
        ("Test cases generated", len(cases)),
        ("Inspection mode", "Read-only DOM inspection"),
        ("Saved input data", "None"),
        ("Warnings", " | ".join(observations.get("warnings", [])) or "None"),
    ]
    for row in rows:
        info.append(list(row))
    _style_sheet(info)
    wb.save(output_path)
    return output_path


def json_safe(value: Any) -> str:
    import json
    return json.dumps(value, ensure_ascii=False)
