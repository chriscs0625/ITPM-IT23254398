# IT3040 – ITPM Assignment 1 – Playwright Automation

## Feature Tested
**Image Format Conversion** – Preview functionality on [https://www.pixelssuite.com/convert-to-png](https://www.pixelssuite.com/convert-to-png)

## Prerequisites
- Python 3.11 or 3.12
- Google Chrome installed

## Installation (one-time)

```bash
pip install -U pip
pip install playwright openpyxl
playwright install
```

## Run the Test

```bash
python image_preview_test.py --url "https://www.pixelssuite.com/convert-to-png" --slow-mo-ms 2000
```

## Output Files
- `execution_results.csv` – Test result record
- `results/preview_pass.png` – Screenshot taken during test

## Project Structure
```
test_automation_ui/
├── image_preview_test.py
├── sample.png
├── execution_results.csv
├── README.md
└── results/
    └── preview_pass.png
```
