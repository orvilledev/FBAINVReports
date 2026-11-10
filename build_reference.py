"""Rebuild reference_schema.json from a workbook: python build_reference.py "path\\to\\file.xlsx" """

import sys

from schema_compare import save_reference

if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    ref = save_reference(sys.argv[1])
    for t in ref["tabs"]:
        print(f"{t['key']:<25} {len(t['columns']):>3} columns   (tab: {t['name']})")
