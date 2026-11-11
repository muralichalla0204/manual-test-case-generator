# Website-to-Automation Test Case Generator

A read-only Python + Playwright website inspector that creates an Excel workbook of
basic, evidence-based manual test cases. It does **not** save or use credentials,
submit forms, click potentially consequential controls, or attempt authentication.

## What it does

1. Accepts a website URL from the command line.
2. Opens the page with Playwright in a headless Chromium browser.
3. Collects observable facts: page title, final URL, headings, links, buttons,
   input fields, textareas, selects, and forms.
4. Generates test cases only from observed page information plus clearly labelled
   generic checks.
5. Writes an Excel workbook with `Test Cases`, `Page Observations`, and `Run Info`.

## Requirements

- Python 3.10+
- Internet access for the target website
- Permission to inspect the target website

## Setup

### Windows (PowerShell)

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m playwright install chromium
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m playwright install chromium
```

## Run

```bash
python main.py https://example.com
```

Choose an output filename:

```bash
python main.py https://example.com --output output/example_test_cases.xlsx
```

Set a navigation timeout (milliseconds):

```bash
python main.py https://example.com --timeout 45000
```

The workbook is saved to the path printed in the terminal.

## Workbook sheets

- **Test Cases**: test ID, scenario, preconditions, steps, expected result,
  priority, type, source evidence, and notes.
- **Page Observations**: the page facts collected by Playwright.
- **Run Info**: requested URL, final URL, title, timestamp, and inspection status.

## Safety and scope

- Use only on websites you own or are authorized to test.
- The inspector performs read-only DOM inspection. It does not click buttons,
  submit forms, upload files, create accounts, purchase items, or delete data.
- It does not bypass logins, CAPTCHAs, access controls, or bot protections.
- It does not infer hidden application behavior. A generated test may need
  human refinement before execution.
- The tool generates an **Excel test-case document**, not executable automation
  scripts. It is intended as a first step toward a later execution framework.

## Project structure

```text
website_automation_test_generator/
├── generator/
│   ├── __init__.py
│   ├── inspector.py
│   └── testcase_builder.py
├── output/
│   └── .gitkeep
├── main.py
├── requirements.txt
└── README.md
```
