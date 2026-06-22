from __future__ import annotations

import argparse
import os
import re
import tempfile
from collections import OrderedDict
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
DCTYPE_NS = "http://purl.org/dc/dcmitype/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

NS = {
    "m": MAIN_NS,
    "r": REL_NS,
    "pr": PKG_REL_NS,
}

for prefix, uri in [
    ("", MAIN_NS),
    ("r", REL_NS),
    ("cp", CORE_NS),
    ("dc", DC_NS),
    ("dcterms", DCTERMS_NS),
    ("dcmitype", DCTYPE_NS),
    ("xsi", XSI_NS),
    ("vt", VT_NS),
]:
    ET.register_namespace(prefix, uri)


CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")
GENERATED_PREFIXES = ("visitas_por_gerente",)


def q(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def column_letter(column_index: int) -> str:
    result = ""
    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def column_number(column_letters: str) -> int:
    result = 0
    for char in column_letters:
        result = result * 26 + ord(char.upper()) - 64
    return result


def cell_column(cell_ref: str) -> str:
    match = CELL_REF_RE.match(cell_ref)
    return match.group(1) if match else cell_ref


def safe_sheet_name(name: str, used_names: set[str]) -> str:
    safe = re.sub(r"[\\/?*\[\]:]", "-", name).strip().strip("'")
    if not safe:
        safe = "Sem gerente"
    safe = safe[:31]

    base = safe
    index = 2
    while safe.upper() in used_names:
        suffix = f" ({index})"
        safe = f"{base[:31 - len(suffix)]}{suffix}"
        index += 1

    used_names.add(safe.upper())
    return safe


def text_from_shared_string(si: ET.Element) -> str:
    return "".join(t.text or "" for t in si.findall(".//m:t", NS))


def load_shared_strings(zip_file: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []

    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    return [text_from_shared_string(si) for si in root.findall("m:si", NS)]


def cell_value(cell: ET.Element | None, shared_strings: list[str]) -> str | int | float | None:
    if cell is None:
        return None

    cell_type = cell.attrib.get("t")

    if cell_type == "inlineStr":
        inline = cell.find("m:is", NS)
        if inline is None:
            return ""
        return "".join(t.text or "" for t in inline.findall(".//m:t", NS))

    value = cell.find("m:v", NS)
    if value is None or value.text is None:
        return None

    if cell_type == "s":
        return shared_strings[int(value.text)]

    if cell_type == "str":
        return value.text

    text = value.text
    try:
        number = float(text)
        return int(number) if number.is_integer() else number
    except ValueError:
        return text


def read_source_rows(input_path: Path, max_columns: int) -> list[list[str | int | float | None]]:
    with ZipFile(input_path, "r") as source_zip:
        shared_strings = load_shared_strings(source_zip)
        sheet = ET.fromstring(source_zip.read("xl/worksheets/sheet1.xml"))
        sheet_data = sheet.find("m:sheetData", NS)
        if sheet_data is None:
            raise ValueError("Nao encontrei os dados da primeira aba da planilha.")

        rows: list[list[str | int | float | None]] = []
        for row in sheet_data.findall("m:row", NS):
            values: list[str | int | float | None] = [None] * max_columns

            for cell in row.findall("m:c", NS):
                ref = cell.attrib.get("r", "")
                col_index = column_number(cell_column(ref)) if ref else 0
                if 1 <= col_index <= max_columns:
                    values[col_index - 1] = cell_value(cell, shared_strings)

            rows.append(values)

    return rows


def is_empty(value: str | int | float | None) -> bool:
    return value is None or str(value).strip() == ""


def newest_input_file(folder: Path, output_name: str) -> Path:
    candidates = [
        path
        for path in folder.glob("*.xlsx")
        if (
            not path.name.startswith("~$")
            and path.name.lower() != output_name.lower()
            and not path.stem.lower().startswith(GENERATED_PREFIXES)
        )
    ]

    if not candidates:
        raise FileNotFoundError(
            "Nenhuma planilha base .xlsx encontrada na pasta. "
            "Arquivos gerados como 'visitas_por_gerente.xlsx' sao ignorados."
        )

    return max(candidates, key=lambda path: path.stat().st_mtime)


def build_content_types(sheet_count: int) -> bytes:
    root = ET.Element(q(CONTENT_NS, "Types"))
    ET.SubElement(
        root,
        q(CONTENT_NS, "Default"),
        {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"},
    )
    ET.SubElement(
        root,
        q(CONTENT_NS, "Default"),
        {"Extension": "xml", "ContentType": "application/xml"},
    )

    overrides = [
        ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            (
                f"/xl/worksheets/sheet{index}.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            )
        )

    for part_name, content_type in overrides:
        ET.SubElement(
            root,
            q(CONTENT_NS, "Override"),
            {"PartName": part_name, "ContentType": content_type},
        )

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_root_rels() -> bytes:
    root = ET.Element(q(PKG_REL_NS, "Relationships"))
    ET.SubElement(
        root,
        q(PKG_REL_NS, "Relationship"),
        {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "Target": "xl/workbook.xml",
        },
    )
    ET.SubElement(
        root,
        q(PKG_REL_NS, "Relationship"),
        {
            "Id": "rId2",
            "Type": "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "Target": "docProps/core.xml",
        },
    )
    ET.SubElement(
        root,
        q(PKG_REL_NS, "Relationship"),
        {
            "Id": "rId3",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
            "Target": "docProps/app.xml",
        },
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_workbook(sheet_names: list[str]) -> bytes:
    workbook = ET.Element(q(MAIN_NS, "workbook"))
    ET.SubElement(workbook, q(MAIN_NS, "workbookPr"), {"defaultThemeVersion": "166925"})
    book_views = ET.SubElement(workbook, q(MAIN_NS, "bookViews"))
    ET.SubElement(book_views, q(MAIN_NS, "workbookView"), {"activeTab": "0"})
    sheets = ET.SubElement(workbook, q(MAIN_NS, "sheets"))

    for index, name in enumerate(sheet_names, start=1):
        ET.SubElement(
            sheets,
            q(MAIN_NS, "sheet"),
            {
                "name": name,
                "sheetId": str(index),
                q(REL_NS, "id"): f"rId{index}",
            },
        )

    ET.SubElement(workbook, q(MAIN_NS, "calcPr"), {"calcId": "0"})
    return ET.tostring(workbook, encoding="utf-8", xml_declaration=True)


def build_workbook_rels(sheet_count: int) -> bytes:
    root = ET.Element(q(PKG_REL_NS, "Relationships"))
    for index in range(1, sheet_count + 1):
        ET.SubElement(
            root,
            q(PKG_REL_NS, "Relationship"),
            {
                "Id": f"rId{index}",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "Target": f"worksheets/sheet{index}.xml",
            },
        )

    ET.SubElement(
        root,
        q(PKG_REL_NS, "Relationship"),
        {
            "Id": f"rId{sheet_count + 1}",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "Target": "styles.xml",
        },
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_styles() -> bytes:
    root = ET.Element(q(MAIN_NS, "styleSheet"))
    num_fmts = ET.SubElement(root, q(MAIN_NS, "numFmts"), {"count": "1"})
    ET.SubElement(num_fmts, q(MAIN_NS, "numFmt"), {"numFmtId": "164", "formatCode": "dd/mm/yyyy"})

    fonts = ET.SubElement(root, q(MAIN_NS, "fonts"), {"count": "2"})
    font_normal = ET.SubElement(fonts, q(MAIN_NS, "font"))
    ET.SubElement(font_normal, q(MAIN_NS, "sz"), {"val": "11"})
    ET.SubElement(font_normal, q(MAIN_NS, "name"), {"val": "Calibri"})
    font_bold = ET.SubElement(fonts, q(MAIN_NS, "font"))
    ET.SubElement(font_bold, q(MAIN_NS, "b"))
    ET.SubElement(font_bold, q(MAIN_NS, "sz"), {"val": "11"})
    ET.SubElement(font_bold, q(MAIN_NS, "name"), {"val": "Calibri"})

    fills = ET.SubElement(root, q(MAIN_NS, "fills"), {"count": "2"})
    ET.SubElement(fills, q(MAIN_NS, "fill")).append(ET.Element(q(MAIN_NS, "patternFill"), {"patternType": "none"}))
    ET.SubElement(fills, q(MAIN_NS, "fill")).append(ET.Element(q(MAIN_NS, "patternFill"), {"patternType": "gray125"}))

    borders = ET.SubElement(root, q(MAIN_NS, "borders"), {"count": "1"})
    border = ET.SubElement(borders, q(MAIN_NS, "border"))
    for side in ("left", "right", "top", "bottom", "diagonal"):
        ET.SubElement(border, q(MAIN_NS, side))

    cell_style_xfs = ET.SubElement(root, q(MAIN_NS, "cellStyleXfs"), {"count": "1"})
    ET.SubElement(cell_style_xfs, q(MAIN_NS, "xf"), {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0"})

    cell_xfs = ET.SubElement(root, q(MAIN_NS, "cellXfs"), {"count": "3"})
    ET.SubElement(cell_xfs, q(MAIN_NS, "xf"), {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"})
    ET.SubElement(cell_xfs, q(MAIN_NS, "xf"), {"numFmtId": "0", "fontId": "1", "fillId": "0", "borderId": "0", "xfId": "0", "applyFont": "1"})
    ET.SubElement(cell_xfs, q(MAIN_NS, "xf"), {"numFmtId": "164", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0", "applyNumberFormat": "1"})

    cell_styles = ET.SubElement(root, q(MAIN_NS, "cellStyles"), {"count": "1"})
    ET.SubElement(cell_styles, q(MAIN_NS, "cellStyle"), {"name": "Normal", "xfId": "0", "builtinId": "0"})
    ET.SubElement(root, q(MAIN_NS, "dxfs"), {"count": "0"})
    ET.SubElement(root, q(MAIN_NS, "tableStyles"), {"count": "0", "defaultTableStyle": "TableStyleMedium2", "defaultPivotStyle": "PivotStyleLight16"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_core_props() -> bytes:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    root = ET.Element(q(CORE_NS, "coreProperties"))
    ET.SubElement(root, q(DC_NS, "creator")).text = "automatizador_visitas"
    ET.SubElement(root, q(CORE_NS, "lastModifiedBy")).text = "automatizador_visitas"
    created = ET.SubElement(root, q(DCTERMS_NS, "created"), {q(XSI_NS, "type"): "dcterms:W3CDTF"})
    created.text = now
    modified = ET.SubElement(root, q(DCTERMS_NS, "modified"), {q(XSI_NS, "type"): "dcterms:W3CDTF"})
    modified.text = now
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_app_props(sheet_names: list[str]) -> bytes:
    ET.register_namespace("", APP_NS)
    ET.register_namespace("vt", VT_NS)

    root = ET.Element(q(APP_NS, "Properties"))
    ET.SubElement(root, q(APP_NS, "Application")).text = "Microsoft Excel"
    ET.SubElement(root, q(APP_NS, "DocSecurity")).text = "0"
    ET.SubElement(root, q(APP_NS, "ScaleCrop")).text = "false"

    heading_pairs = ET.SubElement(root, q(APP_NS, "HeadingPairs"))
    heading_vector = ET.SubElement(heading_pairs, q(VT_NS, "vector"), {"size": "2", "baseType": "variant"})
    variant_label = ET.SubElement(heading_vector, q(VT_NS, "variant"))
    ET.SubElement(variant_label, q(VT_NS, "lpstr")).text = "Planilhas"
    variant_count = ET.SubElement(heading_vector, q(VT_NS, "variant"))
    ET.SubElement(variant_count, q(VT_NS, "i4")).text = str(len(sheet_names))

    titles = ET.SubElement(root, q(APP_NS, "TitlesOfParts"))
    title_vector = ET.SubElement(titles, q(VT_NS, "vector"), {"size": str(len(sheet_names)), "baseType": "lpstr"})
    for name in sheet_names:
        ET.SubElement(title_vector, q(VT_NS, "lpstr")).text = name

    ET.SubElement(root, q(APP_NS, "Company")).text = ""
    ET.SubElement(root, q(APP_NS, "LinksUpToDate")).text = "false"
    ET.SubElement(root, q(APP_NS, "SharedDoc")).text = "false"
    ET.SubElement(root, q(APP_NS, "HyperlinksChanged")).text = "false"
    ET.SubElement(root, q(APP_NS, "AppVersion")).text = "16.0300"
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_sheet(values: list[list[str | int | float | None]], max_columns: int) -> bytes:
    ET.register_namespace("", MAIN_NS)
    ET.register_namespace("r", REL_NS)

    root = ET.Element(q(MAIN_NS, "worksheet"))
    ET.SubElement(root, q(MAIN_NS, "dimension"), {"ref": f"A1:{column_letter(max_columns)}{max(len(values), 1)}"})
    sheet_views = ET.SubElement(root, q(MAIN_NS, "sheetViews"))
    sheet_view = ET.SubElement(sheet_views, q(MAIN_NS, "sheetView"), {"workbookViewId": "0"})
    ET.SubElement(sheet_view, q(MAIN_NS, "selection"), {"activeCell": "A1", "sqref": "A1"})
    ET.SubElement(root, q(MAIN_NS, "sheetFormatPr"), {"defaultRowHeight": "15"})

    cols = ET.SubElement(root, q(MAIN_NS, "cols"))
    widths = [30, 30, 30, 32, 30, 56, 8, 12, 22, 16, 12, 14, 18, 8]
    for index in range(1, max_columns + 1):
        width = widths[index - 1] if index <= len(widths) else 12
        ET.SubElement(cols, q(MAIN_NS, "col"), {"min": str(index), "max": str(index), "width": str(width), "customWidth": "1"})

    sheet_data = ET.SubElement(root, q(MAIN_NS, "sheetData"))
    date_columns = {10, 11, 12}

    for row_number, row_values in enumerate(values, start=1):
        row = ET.SubElement(sheet_data, q(MAIN_NS, "row"), {"r": str(row_number)})

        for col_index in range(1, max_columns + 1):
            value = row_values[col_index - 1] if col_index <= len(row_values) else None
            if is_empty(value):
                continue

            cell_ref = f"{column_letter(col_index)}{row_number}"
            attrs = {"r": cell_ref}
            if row_number == 1:
                attrs["s"] = "1"
            elif col_index in date_columns and isinstance(value, (int, float)):
                attrs["s"] = "2"

            cell = ET.SubElement(row, q(MAIN_NS, "c"), attrs)

            if isinstance(value, (int, float)):
                ET.SubElement(cell, q(MAIN_NS, "v")).text = str(value)
            else:
                cell.attrib["t"] = "inlineStr"
                inline = ET.SubElement(cell, q(MAIN_NS, "is"))
                text = ET.SubElement(inline, q(MAIN_NS, "t"))
                text.text = str(value)

    ET.SubElement(root, q(MAIN_NS, "pageMargins"), {"left": "0.7", "right": "0.7", "top": "0.75", "bottom": "0.75", "header": "0.3", "footer": "0.3"})
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def next_available_output_path(output_path: Path) -> Path:
    for index in range(2, 100):
        candidate = output_path.with_name(f"{output_path.stem}_{index}{output_path.suffix}")
        if not candidate.exists():
            return candidate

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_path.with_name(f"{output_path.stem}_{timestamp}{output_path.suffix}")


def write_workbook(output_path: Path, sheet_data: OrderedDict[str, list[list[str | int | float | None]]], max_columns: int) -> Path:
    sheet_names = list(sheet_data.keys())
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=output_path.stem + "_", suffix=".xlsx", dir=output_path.parent)
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        with ZipFile(temp_path, "w", ZIP_DEFLATED) as output_zip:
            output_zip.writestr("[Content_Types].xml", build_content_types(len(sheet_names)))
            output_zip.writestr("_rels/.rels", build_root_rels())
            output_zip.writestr("xl/workbook.xml", build_workbook(sheet_names))
            output_zip.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels(len(sheet_names)))
            output_zip.writestr("xl/styles.xml", build_styles())
            output_zip.writestr("docProps/core.xml", build_core_props())
            output_zip.writestr("docProps/app.xml", build_app_props(sheet_names))

            for index, rows in enumerate(sheet_data.values(), start=1):
                output_zip.writestr(f"xl/worksheets/sheet{index}.xml", build_sheet(rows, max_columns))

        try:
            os.replace(temp_path, output_path)
            return output_path
        except PermissionError:
            fallback_path = next_available_output_path(output_path)
            os.replace(temp_path, fallback_path)
            print(f"Aviso: '{output_path.name}' esta aberto no Excel.")
            print(f"Salvei uma nova copia em: {fallback_path.resolve()}")
            return fallback_path
    finally:
        if temp_path.exists():
            temp_path.unlink()


def create_manager_workbook(
    input_path: Path,
    output_path: Path,
    header_row_number: int,
    manager_column_number: int,
    base_sheet_name: str,
    max_columns: int,
) -> tuple[list[str], dict[str, int]]:
    input_path = input_path.resolve()
    output_path = output_path.resolve()

    if input_path == output_path:
        raise ValueError("O arquivo de entrada e saida nao podem ser o mesmo.")

    rows = read_source_rows(input_path, max_columns)
    if len(rows) < header_row_number:
        raise ValueError("A planilha nao tem cabecalho suficiente.")

    header = rows[header_row_number - 1]
    data_rows = rows[header_row_number:]
    manager_index = manager_column_number - 1

    rows_by_manager: OrderedDict[str, list[list[str | int | float | None]]] = OrderedDict()
    base_rows = [header]

    for row in data_rows:
        manager = "" if manager_index >= len(row) or row[manager_index] is None else str(row[manager_index]).strip()
        if not manager:
            continue

        base_rows.append(row)
        rows_by_manager.setdefault(manager, []).append(row)

    used_sheet_names = {base_sheet_name.upper()}
    sheet_data: OrderedDict[str, list[list[str | int | float | None]]] = OrderedDict()
    sheet_data[base_sheet_name] = base_rows

    manager_counts: dict[str, int] = {}
    for manager, manager_rows in rows_by_manager.items():
        sheet_name = safe_sheet_name(manager, used_sheet_names)
        sheet_data[sheet_name] = [header] + manager_rows
        manager_counts[manager] = len(manager_rows)

    saved_path = write_workbook(output_path, sheet_data, max_columns)
    return list(sheet_data.keys()), manager_counts, saved_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Cria uma aba por gerente usando a coluna D da planilha base.")
    parser.add_argument("-i", "--input", type=Path, default=None, help="Arquivo .xlsx de entrada.")
    parser.add_argument("-o", "--output", type=Path, default=None, help="Arquivo .xlsx de saida. Padrao: visitas_por_gerente.xlsx.")
    parser.add_argument("--header-row", type=int, default=1)
    parser.add_argument("--manager-column", type=int, default=4)
    parser.add_argument("--base-sheet-name", default="Base Total")
    parser.add_argument("--max-columns", type=int, default=14)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    script_folder = Path(__file__).resolve().parent
    output_path = args.output or script_folder / "visitas_por_gerente.xlsx"
    input_path = args.input or newest_input_file(script_folder, output_path.name)

    print(f"Planilha base usada: {input_path.resolve()}")
    sheet_names, manager_counts, saved_path = create_manager_workbook(
        input_path=input_path,
        output_path=output_path,
        header_row_number=args.header_row,
        manager_column_number=args.manager_column,
        base_sheet_name=args.base_sheet_name,
        max_columns=args.max_columns,
    )

    print(f"Arquivo criado: {saved_path.resolve()}")
    print(f"Gerentes encontrados: {len(manager_counts)}")
    print("Abas:")
    for name in sheet_names:
        print(f"- {name}")
    print("Linhas copiadas por gerente:")
    for manager, line_count in manager_counts.items():
        print(f"- {manager}: {line_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
