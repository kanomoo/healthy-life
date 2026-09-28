import sys, re

sys.stdout.reconfigure(encoding='utf-8')

with open('scripts/build_full_report.py', 'r', encoding='utf-8') as f:
    text = f.read()

# Find all table definitions:
table_defs = re.findall(r'(headers_\w+\s*=\s*\[.*?\]\s*\n\s*data_\w+\s*=\s*\[.*?\]\s*\n\s*col_w_\w+\s*=\s*\[.*?\])', text, re.DOTALL)
print(f"Total tables found: {len(table_defs)}")
for i, td in enumerate(table_defs):
    name_m = re.search(r'headers_(\w+)', td)
    name = name_m.group(1) if name_m else str(i)
    print(f"\n--- Table {name} ---")
    lines = td.splitlines()
    print(lines[0]) # headers
    col_w_line = [l for l in lines if 'col_w_' in l]
    if col_w_line:
        print(col_w_line[0])
