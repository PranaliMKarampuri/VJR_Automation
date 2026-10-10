import json
from openpyxl.utils import get_column_letter

d = json.load(open("reports/pavillon_c_vjr_comparison/workbook_contents.json", encoding="utf-8"))
for sh in ["TAKE OFF", "SPIRALE", "ROND ET FLEX"]:
    print("###", sh)
    for label in ["manual", "generated"]:
        print("--", label)
        count = 0
        for record in d[label]["sheets"][sh]["populated_rows"]:
            r = record["row"]
            cells = [f"{get_column_letter(i+1)}={v!r}" for i,v in enumerate(record["values"]) if v is not None]
            print(r, " ".join(cells[:25]))
            count += 1
            if count >= 30: break
