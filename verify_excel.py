import openpyxl
from collections import Counter
from pathlib import Path

file_path = Path("Manual Test Cases for Option 2.xlsx")

# Load the workbook
wb = openpyxl.load_workbook(file_path)
ws = wb["Manual Test Cases"]

# Read all data rows (skipping the header)
data_rows = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is not None:  # Ensure the row isn't empty
        data_rows.append((row[0], row[1]))

print(f"Total data rows (excluding header): {len(data_rows)}\n")
print(f"{'TC ID':<12} | {'Application Feature Tested'}")
print("-" * 50)

# Print each TC ID and Feature, and count occurrences
feature_counts = Counter()
for tc_id, feature in data_rows:
    print(f"{tc_id:<12} | {feature}")
    feature_counts[feature] += 1

print("\n--- Feature Coverage ---")
for feature, count in feature_counts.items():
    print(f"- {feature}: {count} test cases")

# Verification Checks
print("\n--- Final Verification ---")
print(f"1. Exactly 36 rows? {'Yes' if len(data_rows) == 36 else 'No (Count is ' + str(len(data_rows)) + ')'}")
print(f"2. Total unique features covered: {len(feature_counts)}")
print(f"3. Are all 10 features covered? {'Yes' if len(feature_counts) == 10 else 'No'}")
