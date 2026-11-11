from pathlib import Path
from playwright.sync_api import sync_playwright

from generator.inspector import inspect_website, normalize_url
from generator.testcase_builder import write_workbook


def main():
    print("=" * 60)
    print(" WEBSITE-TO-AUTOMATION TEST CASE GENERATOR")
    print("=" * 60)

    raw_url = input(
        "\nEnter website URL to inspect: "
    ).strip()

    try:
        url = normalize_url(raw_url)
    except ValueError as error:
        print(f"Invalid URL: {error}")
        return

    print("\nInspecting website...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        try:
            observations = inspect_website(browser, url)
        finally:
            browser.close()

    output = write_workbook(
        observations,
        "output/generated_test_cases.xlsx"
    )

    print("\nTest case generation completed!")
    print(f"Excel file created: {Path(output).resolve()}")


if __name__ == "__main__":
    main()