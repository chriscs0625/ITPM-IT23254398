from playwright.sync_api import sync_playwright
from pathlib import Path
import argparse
import base64
import time
import sys
import csv
import os

DEFAULT_URL = "https://www.pixelssuite.com/convert-to-png"
DEFAULT_TIMEOUT_MS = 60000
DEFAULT_SLOW_MO_MS = 0

PNG_IX_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAIAAAAiOjnJAAADhElEQVR4nO3dMUudVxzA4dfGJZlCh0ySbyC4KgrXQQcHt2QWTJdOfgSnZMuWDhayK0qWfATBzdElo4sYLJmCOHWQlmJtaUt+vXp9nunAfV/4X/jxnrOdqVfvLgb41r4b9wBMJmGREBYJYZEQFglhkRAWCWGREBYJYZEQFglhkRAWCWGREBYJYZEQFglhkRAWCWGREBYJYZEQFglhkRAWCWGREBYJYZEQFglhkRAWCWGREBYJYZEQFglhkRAWCWGREBYJYZEQFglhkRAWCWGREBYJYZEQFglhkRAWmelxD3C7n3/8ftwj3Cc//PTLuEe4yReLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi8QEhvXhw4fRaDQajaanp68X+/v7T548Gf3m7du3wzAcHx+vrq4uLy+vrKycnp7e+ta4/8o9NvXq3cW4Z7jFN7ny5OnTp1++fPnz+trc3NzHjx9nZmYODg729vZ2d3f/6sm77w5eeXJH79L5H5yfn19eXg7DsL6+/uzZs3GPM2kmcCv8h16/fr20tLS5uXl4eLi0tDTucSbNQwnr6urq9zPW0dHRMAwbGxsnJyeLi4tbW1vb29vjHnDSPNAz1ufPnz99+rSwsHC9np2dPTs7u/XJe+EOnrEeyhfrhqmpqZcvX56eng7DcHFx8fz583FPNGkeyuH9eiu8Xs/Pz79582ZnZ+fFixePHz9+9OjR+/fvxzrdBJrksP64o339+vXGr2tra2tra3//Fv/ZA90KqQmLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLhLBICIuEsEgIi4SwSAiLxB29uvcOXsvOv+KLRUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCSERUJYJIRFQlgkhEVCWCR+BV/FiqJF1yrFAAAAAElFTkSuQmCC"

def configure_stdout():
    sys.stdout.reconfigure(encoding='utf-8') if hasattr(sys.stdout, 'reconfigure') else None

def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default=DEFAULT_URL)
    parser.add_argument("--slow-mo-ms", type=int, default=DEFAULT_SLOW_MO_MS)
    return parser.parse_args()

def create_sample_png(path: Path):
    path.write_bytes(base64.b64decode(PNG_IX_BASE64))

def run_test(url: str, slow_mo: int):
    configure_stdout()

    script_dir = Path(__file__).parent
    sample_path = script_dir / "sample.png"
    results_dir = script_dir / "results"
    results_dir.mkdir(exist_ok=True)

    if not sample_path.exists():
        create_sample_png(sample_path)

    csv_path = script_dir / "execution_results.csv"
    screenshot_path = results_dir / "preview_pass.png"

    preview_detected = False
    status = "FAIL"
    screenshot_str = str(screenshot_path)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=slow_mo)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT_MS)

        try:
            page.goto(url, wait_until="domcontentloaded")
            time.sleep(2)

            # Try to find file input - pixelssuite uses hidden input triggered by button
            file_input = page.locator("input[type='file']").first
            file_input.set_input_files(str(sample_path))
            time.sleep(3)

            # Wait for preview image to appear
            # pixelssuite shows preview in an img tag after upload
            preview_selectors = [
                "img[src^='blob:']",
                "img[src^='data:']",
                ".preview img",
                "#preview img",
                ".image-preview img",
                "canvas",
                ".result-image img",
                "img.preview",
                "img[alt*='preview' i]",
                "img[alt*='image' i]",
            ]

            for sel in preview_selectors:
                try:
                    el = page.locator(sel).first
                    el.wait_for(state="visible", timeout=8000)
                    preview_detected = True
                    status = "PASS"
                    print(f"Preview detected with selector: {sel}")
                    break
                except Exception:
                    continue

            # Fallback: check if any img appeared that wasn't there before
            if not preview_detected:
                try:
                    page.wait_for_selector("img", timeout=5000)
                    imgs = page.locator("img").all()
                    for img in imgs:
                        src = img.get_attribute("src") or ""
                        if "blob:" in src or "data:" in src or "sample" in src.lower():
                            preview_detected = True
                            status = "PASS"
                            break
                except Exception:
                    pass

            page.screenshot(path=str(screenshot_path))

        except Exception as e:
            print(f"Error during test: {e}")
            try:
                page.screenshot(path=str(screenshot_path))
            except Exception:
                pass
        finally:
            browser.close()

    # Write CSV
    file_exists = csv_path.exists()
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["file_type", "file_path", "preview_detected", "status", "screenshot"])
        writer.writerow(["PNG", str(sample_path), str(preview_detected).upper(), status, str(screenshot_path)])

    print("\n======= TEST RESULT =======")
    print(f"PNG file    : {sample_path}")
    print(f"Preview detected: {preview_detected}")
    print(f"Status      : {status}")
    print(f"Screenshot  : {screenshot_path}")
    print(f"CSV         : {csv_path}")

if __name__ == "__main__":
    args = get_args()
    run_test(args.url, args.slow_mo_ms)
