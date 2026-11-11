from __future__ import annotations

from typing import Any
from urllib.parse import urlparse

from playwright.sync_api import Browser, TimeoutError as PlaywrightTimeoutError


def normalize_url(raw_url: str) -> str:
    """Add https:// when a scheme is omitted and validate the URL shape."""
    raw_url = raw_url.strip()
    if not raw_url:
        raise ValueError("URL cannot be empty.")
    if "://" not in raw_url:
        raw_url = "https://" + raw_url

    parsed = urlparse(raw_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Please provide a valid http:// or https:// URL.")
    return raw_url


def inspect_website(browser: Browser, url: str, timeout_ms: int = 30000) -> dict[str, Any]:
    """
    Inspect a page without clicking controls or submitting forms.
    Only public, currently rendered DOM facts are collected.
    """
    context = browser.new_context(
        viewport={"width": 1440, "height": 900},
        ignore_https_errors=False,
    )
    page = context.new_page()
    page.set_default_timeout(timeout_ms)

    result: dict[str, Any] = {
        "requested_url": url,
        "final_url": url,
        "title": "",
        "status": None,
        "headings": [],
        "links": [],
        "buttons": [],
        "inputs": [],
        "textareas": [],
        "selects": [],
        "forms": [],
        "warnings": [],
    }

    try:
        response = page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        result["status"] = response.status if response else None
        try:
            page.wait_for_load_state("networkidle", timeout=min(timeout_ms, 7000))
        except PlaywrightTimeoutError:
            result["warnings"].append(
                "Network did not become idle within the short wait; inspection continued."
            )

        result["final_url"] = page.url
        result["title"] = page.title()

        # Extract facts from the rendered DOM. No controls are activated.
        data = page.evaluate("""() => {
          const clean = (s) => (s || '').replace(/\\\\s+/g, ' ').trim();
          const labelFor = (el) => {
            const aria = el.getAttribute('aria-label');
            if (aria) return clean(aria);
            if (el.id) {
              const label = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
              if (label) return clean(label.innerText);
            }
            const wrapping = el.closest('label');
            if (wrapping) return clean(wrapping.innerText);
            return clean(el.getAttribute('placeholder') || el.getAttribute('name') || '');
          };
          const visible = (el) => {
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' &&
                   s.display !== 'none';
          };
          const unique = (arr) => arr.filter((x, i) =>
            arr.findIndex(y => JSON.stringify(y) === JSON.stringify(x)) === i);

          const headings = [...document.querySelectorAll('h1,h2,h3,h4,h5,h6')]
            .filter(visible).map(el => ({
              tag: el.tagName.toLowerCase(), text: clean(el.innerText)
            })).filter(x => x.text).slice(0, 80);

          const links = [...document.querySelectorAll('a')]
            .filter(visible).map(el => ({
              text: clean(el.innerText || el.getAttribute('aria-label')),
              href: el.href || '',
              target: el.getAttribute('target') || ''
            })).filter(x => x.text || x.href).slice(0, 150);

          const buttons = [...document.querySelectorAll(
            'button, input[type="button"], input[type="submit"], [role="button"]'
          )].filter(visible).map(el => ({
            text: clean(el.innerText || el.value || el.getAttribute('aria-label')),
            type: el.getAttribute('type') || el.tagName.toLowerCase(),
            disabled: !!el.disabled
          })).slice(0, 100);

          const inputs = [...document.querySelectorAll('input')]
            .filter(visible).map(el => ({
              label: labelFor(el),
              type: (el.getAttribute('type') || 'text').toLowerCase(),
              name: el.getAttribute('name') || '',
              placeholder: el.getAttribute('placeholder') || '',
              required: !!el.required,
              disabled: !!el.disabled
            })).filter(x => !['hidden','submit','button','reset','image'].includes(x.type))
            .slice(0, 100);

          const textareas = [...document.querySelectorAll('textarea')]
            .filter(visible).map(el => ({
              label: labelFor(el), name: el.getAttribute('name') || '',
              placeholder: el.getAttribute('placeholder') || '',
              required: !!el.required, disabled: !!el.disabled
            })).slice(0, 50);

          const selects = [...document.querySelectorAll('select')]
            .filter(visible).map(el => ({
              label: labelFor(el), name: el.getAttribute('name') || '',
              required: !!el.required, disabled: !!el.disabled,
              option_count: el.options.length
            })).slice(0, 50);

          const forms = [...document.querySelectorAll('form')]
            .filter(visible).map(el => ({
              name: el.getAttribute('name') || '',
              id: el.id || '',
              method: (el.getAttribute('method') || 'get').toLowerCase(),
              field_count: el.querySelectorAll('input,textarea,select').length
            })).slice(0, 50);

          return {
            headings, links: unique(links), buttons, inputs,
            textareas, selects, forms,
            document_ready_state: document.readyState
          };
        }""")
        result.update(data)

        if response and response.status >= 400:
            result["warnings"].append(
                f"Server returned HTTP {response.status}; page content may be an error page."
            )
        if not result["headings"]:
            result["warnings"].append("No visible heading elements (h1-h6) were detected.")
    except Exception as exc:
        result["warnings"].append(f"Inspection error: {type(exc).__name__}: {exc}")
    finally:
        context.close()

    return result
