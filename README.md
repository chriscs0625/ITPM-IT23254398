# Functional and Usability Testing – Playwright Automation

## Objective
The main objective of this project is to test the preview functionality of the [Pixels Suite](https://www.pixelssuite.com/) website using Playwright. The assessment evaluates whether the user interface is clear, consistent, and behaves as expected. Note that backend API testing, performance testing, scalability testing, and security testing are strictly outside the scope of this assignment.

## Automated Test Scenario
This test script automates a specific positive test scenario: uploading a valid PNG image file and verifying that the image is successfully detected and displayed in the application's preview section.

## Prerequisites
* Python 3.11 or newer
* `pip` (Python package manager)
* Playwright

## Setup Instructions
Use the following commands to install the required Python packages and web browsers for Playwright:

```bash
pip install -U pip
pip install playwright openpyxl
playwright install
```

## How to Run the Test
To execute the automated UI test, run the following command in your terminal:

```bash
python image_preview_test.py --url "https://www.pixelssuite.com/convert-to-png" --slow-mo-ms 2000
```

## Output Files
After execution, the script generates the following artifacts:
* `execution_results.csv`: Contains the structured test execution result logging the success or failure status.
* `results/preview_pass.png`: A full-page screenshot capturing the exact state of the browser, serving as visual evidence of the test.

## Expected Result
The uploaded PNG image should be correctly displayed internally in the preview section.

## Actual Result
Preview detected successfully (**PASS**).

## Project Structure
```text
.
├── results/                               # Contains screenshots of test executions
│   └── preview_pass.png                   # Evidence of successful UI preview detection
├── execution_results.csv                  # Automated logging of test execution statuses
├── image_preview_test.py                  # The main Playwright Python automation script
├── Manual Test Cases for Option 2.xlsx    # Full suite of 36 manual UI test cases
├── sample.png                             # Target image file used for test automation
└── README.md                              # Professional project implementation instructions
```

## Notes
This project completes and fulfills the assignment requirement to effectively automate one test scenario and accurately record the execution results automatically in a CSV file formats.