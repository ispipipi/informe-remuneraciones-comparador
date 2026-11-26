"""Run the accumulated-book regression against the November fixture."""
import json
import runpy
import subprocess
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

root = Path(__file__).resolve().parents[1]
module = runpy.run_path(str(root / "scripts" / "generate_remuneraciones_html.py"))
source = root / "data" / "avanza_acumulado_2026-11.xlsx"
workbook = load_workbook(source, read_only=True, data_only=True)
try:
    aoa = [list(row) for row in workbook["Libro de remu CE"].values]
finally:
    workbook.close()
headers = aoa[5]
index = {header: i for i, header in enumerate(headers)}
totals = defaultdict(lambda: {"haberes": 0, "liquido": 0})
for row in aoa[6:]:
    if not row or not row[0]:
        continue
    month = module["period_id"](row[index["Liquidación - Período"]])
    totals[month]["haberes"] += row[index["Liquidación - Total Haberes"]] or 0
    totals[month]["liquido"] += row[index["Liquidación - Sueldo Líquido"]] or 0
details, _, _, _ = module["avanza_details"](source)
fixture = {
    "template": module["HTML_TEMPLATE"],
    "aoa": aoa,
    "totals": totals,
    "pythonRows": details,
}
subprocess.run(
    ["node", str(root / "tests" / "avanza_accumulated.cjs")],
    input=json.dumps(fixture, ensure_ascii=False, default=str),
    text=True,
    check=True,
    cwd=root,
)
