from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
PREVIEW_DATA_PATH = ROOT / "control-policial-data.js"
REACT_DATA_PATH = ROOT / "src" / "controlData.js"
SHAREPOINT_EXCEL_PATH = Path(
    r"C:\Users\acordova\OneDrive - EMOVA MOVILIDAD S.A\COMPARTIDO CMI - COMPARTIDO CMI\VARIAS\02 - PLANILLA DE CONTROL DE PERSONAL POLICIAL\2026\05 - CONTROL POLICIAL - MAYO.xlsx"
)

HEADERS = [
    "",
    "dni",
    "nombre",
    "turno_excel",
    "dniCubre",
    "nombreCubre",
    "linea",
    "estacion",
    "ingreso",
    "egreso",
    "fiscalizado",
    "novedad",
    "controlCmi",
]


def find_latest_excel() -> Path:
    if SHAREPOINT_EXCEL_PATH.exists():
        return SHAREPOINT_EXCEL_PATH

    candidates = [
        path
        for path in ROOT.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls"}
        and not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError("No se encontraron archivos Excel .xlsx/.xls en la carpeta del proyecto.")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def clean(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def date_from_sheet(sheet_name: str) -> str:
    match = re.search(r"(\d{2})-(\d{2})-(\d{4})", sheet_name)
    if not match:
        return sheet_name.strip()
    return f"{match.group(3)}-{match.group(2)}-{match.group(1)}"


def format_line(value: str) -> str:
    line = clean(value).upper().replace("Í", "I")
    if line in list("ABCDEH"):
        return f"Línea {line}"
    return line or "Sin línea"


def normalize_shift(turno: str, ingreso: str) -> str:
    text = f"{clean(turno)} {clean(ingreso)}".lower()
    match = re.search(r"(\d{1,2}):", text)
    hour = int(match.group(1)) if match else None
    if "16:30" in text or "tarde" in text or (hour is not None and hour >= 14):
        return "16:30 A 00:00 - Turno Tarde"
    return "05:30 A 13:30 - Turno Mañana"


def status(fiscalizado: str, egreso: str, novedad: str, control_cmi: str) -> str:
    text = " ".join([clean(fiscalizado), clean(egreso), clean(novedad), clean(control_cmi)]).lower()
    if "no encontro" in text or "no encontró" in text or "ausente" in text:
        return "Crítico"
    if clean(fiscalizado).upper() == "NO" or "no fiscalizo" in text or "no fiscalizó" in text:
        return "Observación"
    if clean(egreso).upper() == "S/I":
        return "Observación"
    return "Operativo"


def build_records(excel_path: Path) -> list[dict]:
    workbook = pd.ExcelFile(excel_path)
    records: list[dict] = []

    for sheet_name in workbook.sheet_names:
        frame = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=object)
        day = sheet_name.strip()
        date = date_from_sheet(sheet_name)

        for index in range(7, len(frame)):
            row = [clean(value) for value in frame.iloc[index].tolist()[:13]]
            row += [""] * (13 - len(row))
            data = dict(zip(HEADERS, row))

            useful_fields = ["dni", "nombre", "linea", "estacion", "ingreso", "egreso", "fiscalizado", "novedad", "controlCmi"]
            if not any(data.get(field) for field in useful_fields):
                continue
            if data["dni"].upper() == "DNI" or "APELLIDO" in data["nombre"].upper():
                continue

            records.append(
                {
                    "id": f"{day}-{index + 1}",
                    "dia": day,
                    "fecha": date,
                    "hora": data["ingreso"],
                    "dni": data["dni"],
                    "nombre": data["nombre"],
                    "dniCubre": data["dniCubre"],
                    "nombreCubre": data["nombreCubre"],
                    "linea": format_line(data["linea"]),
                    "estacion": data["estacion"],
                    "turno": normalize_shift(data["turno_excel"], data["ingreso"]),
                    "turnoExcel": data["turno_excel"],
                    "ingreso": data["ingreso"],
                    "egreso": data["egreso"],
                    "fiscalizado": data["fiscalizado"] or "-",
                    "dependencia": "Policía de la Ciudad",
                    "dotacion": 2 if data["nombreCubre"] and data["nombreCubre"] != "-" else 1,
                    "puesto": data["nombreCubre"] if data["nombreCubre"] and data["nombreCubre"] != "-" else "Titular",
                    "estado": status(data["fiscalizado"], data["egreso"], data["novedad"], data["controlCmi"]),
                    "controlCmi": data["controlCmi"],
                    "novedad": data["novedad"],
                }
            )

    return records


def main() -> None:
    excel_path = find_latest_excel()
    records = build_records(excel_path)
    payload = json.dumps(records, ensure_ascii=False, indent=2)
    source_payload = json.dumps(excel_path.name, ensure_ascii=False)
    PREVIEW_DATA_PATH.write_text(
        f"window.CONTROL_POLICIAL_SOURCE = {source_payload};\nwindow.CONTROL_POLICIAL_RECORDS = {payload};\n",
        encoding="utf-8",
    )
    REACT_DATA_PATH.write_text(
        f"export const controlSource = {source_payload};\nexport const controlRecords = {payload};\n",
        encoding="utf-8",
    )
    print(
        json.dumps(
            {
                "source": excel_path.name,
                "records": len(records),
                "output": [str(PREVIEW_DATA_PATH), str(REACT_DATA_PATH)],
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
