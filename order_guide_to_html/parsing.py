from __future__ import annotations

from collections import OrderedDict
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import openpyxl
import re

from .utils import CODE_IN_PARENS_RE, FILLER_TOKENS, FOOTNOTE_LINE_RE, STATUS_LABELS, TRAILING_DIGITS_RE, normalize_text, unique_preserve_order
from .models import ColorExteriorRow, ColorInteriorRow, ColorSheet, EngineAxleEntry, EngineAxleItem, GCWRRecord, MatrixRow, MatrixSheet, SpecCell, SpecColumn, TraileringRecord, TrimDef, WorkbookData


def parse_filename_metadata(path: Path) -> Tuple[str, str, str, str]:
    stem = path.stem
    stem = re.sub(r"\s+Export$", "", stem, flags=re.I)
    stem = re.sub(r"\s+(Retail and Fleet|Retail|Fleet)$", "", stem, flags=re.I)
    parts = stem.split()
    if not parts or not re.fullmatch(r"\d{4}", parts[0]):
        raise ValueError(f"Cannot determine year from filename: {path.name}")
    year = parts[0]
    if len(parts) < 3:
        raise ValueError(f"Cannot determine make/model from filename: {path.name}")
    make = parts[1]
    model_tokens = [p for p in parts[2:] if p not in FILLER_TOKENS]
    if not model_tokens:
        raise ValueError(f"Cannot determine model from filename: {path.name}")
    model = " ".join(model_tokens)
    vehicle_name = f"{year} {make} {model}"
    return year, make, model, vehicle_name

def parse_trim_header(value: object) -> Optional[TrimDef]:
    text = normalize_text(value)
    if not text:
        return None
    lines = [normalize_text(x) for x in text.split("\n") if normalize_text(x)]
    if not lines:
        return None
    if len(lines) == 1:
        return TrimDef(name=lines[0], code=lines[0], raw_header=text)
    return TrimDef(name=" ".join(lines[:-1]), code=lines[-1], raw_header=text)

def find_matrix_header_row(ws) -> Optional[int]:
    for r in range(1, min(ws.max_row, 10) + 1):
        values = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "Description" in values:
            return r
    return None

def parse_footnote_map(text: str) -> Dict[str, str]:
    notes: Dict[str, str] = {}
    for line in normalize_text(text).split("\n"):
        m = FOOTNOTE_LINE_RE.match(line)
        if m:
            notes[m.group(1)] = normalize_text(m.group(2))
    return notes

def split_main_notes_and_bullets(text: str) -> Tuple[str, Dict[str, str], List[str]]:
    lines = [normalize_text(line) for line in normalize_text(text).split("\n") if normalize_text(line)]
    main_lines: List[str] = []
    notes: Dict[str, str] = {}
    bullets: List[str] = []
    in_notes = False
    for line in lines:
        m = FOOTNOTE_LINE_RE.match(line)
        if m:
            notes[m.group(1)] = normalize_text(m.group(2))
            in_notes = True
            continue
        if line.startswith("•"):
            bullets.append(normalize_text(line.lstrip("•").strip()))
            in_notes = True
            continue
        if not in_notes:
            main_lines.append(line)
        else:
            bullets.append(line)
    main = normalize_text(" ".join(main_lines))
    return main, notes, bullets

def parse_status_value(raw: str, row_notes: Dict[str, str], sheet_notes: Dict[str, str]) -> Tuple[str, str, List[str]]:
    raw = normalize_text(raw)
    if not raw:
        return "", "", []
    m = re.match(r"^(--|[A-Z]+|[■□*]+)(.*)$", raw)
    if m:
        code = m.group(1)
        suffix = m.group(2)
    else:
        code = raw
        suffix = ""
    label = STATUS_LABELS.get(code, code)
    note_ids = re.findall(r"\d+", suffix)
    notes: List[str] = []
    for note_id in note_ids:
        if note_id in row_notes:
            notes.append(row_notes[note_id])
        elif note_id in sheet_notes:
            notes.append(sheet_notes[note_id])
    return code, label, unique_preserve_order(notes)

def parse_matrix_sheet(ws, trim_defs: Optional[List[TrimDef]] = None) -> Optional[MatrixSheet]:
    header_row = find_matrix_header_row(ws)
    if header_row is None:
        return None

    headers = [normalize_text(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
    description_col = headers.index("Description") + 1

    legend_text_parts: List[str] = []
    sheet_footnotes: Dict[str, str] = {}
    for r in range(1, header_row):
        row_texts = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        nonempty = [t for t in row_texts if t]
        if nonempty:
            legend_text_parts.extend(nonempty)
        for item in nonempty:
            sheet_footnotes.update(parse_footnote_map(item))

    inferred_trim_defs: List[TrimDef] = []
    c = description_col + 1
    while c <= ws.max_column:
        raw = normalize_text(ws.cell(header_row, c).value)
        if not raw:
            break
        parsed = parse_trim_header(raw)
        if parsed:
            inferred_trim_defs.append(parsed)
        c += 1
    active_trim_defs = trim_defs or inferred_trim_defs
    trim_cols = list(range(description_col + 1, description_col + 1 + len(active_trim_defs)))

    rows: List[MatrixRow] = []
    current_group: Optional[str] = None
    for r in range(header_row + 1, ws.max_row + 1):
        meta = [normalize_text(ws.cell(r, c).value) for c in range(1, description_col + 1)]
        status_values = [normalize_text(ws.cell(r, c).value) for c in trim_cols]
        if not any(meta) and not any(status_values):
            continue

        if not any(status_values):
            nonempty = [m for m in meta if m]
            for item in nonempty:
                sheet_footnotes.update(parse_footnote_map(item))
            if nonempty and not any(FOOTNOTE_LINE_RE.match(line) for item in nonempty for line in item.split("\n")):
                current_group = nonempty[0]
            continue

        description_raw = meta[-1]
        if not description_raw:
            continue

        option_code = meta[0] or None
        ref_code = meta[1] or None if len(meta) > 1 else None
        aux_meta = [x for x in meta[2:-1] if x]
        main, inline_notes, bullet_notes = split_main_notes_and_bullets(description_raw)

        row = MatrixRow(
            sheet_name=ws.title,
            row_group=current_group,
            option_code=option_code,
            ref_code=ref_code,
            aux_meta=aux_meta,
            description_raw=description_raw,
            description_main=main or description_raw,
            inline_footnotes=inline_notes,
            bullet_notes=bullet_notes,
            status_by_trim={
                active_trim_defs[idx].key: status_values[idx]
                for idx in range(min(len(active_trim_defs), len(status_values)))
                if normalize_text(status_values[idx])
            },
        )
        rows.append(row)

    legend_text = "\n".join(unique_preserve_order(legend_text_parts))
    return MatrixSheet(
        name=ws.title,
        legend_text=legend_text,
        trim_defs=active_trim_defs,
        footnotes=sheet_footnotes,
        rows=rows,
    )

def parse_color_sheet(ws) -> ColorSheet:
    footnotes: Dict[str, str] = {}
    bullet_notes: List[str] = []
    heading_lines: List[str] = []
    interior_rows: List[ColorInteriorRow] = []
    exterior_rows: List[ColorExteriorRow] = []

    for r in range(1, ws.max_row + 1):
        row_values = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        first = row_values[0] if row_values else ""
        if r <= 3 and any(row_values):
            heading_lines.extend([v for v in row_values if v])
        for value in row_values:
            if value:
                footnotes.update(parse_footnote_map(value))
                for line in normalize_text(value).split("\n"):
                    if line.startswith("•"):
                        bullet_notes.append(normalize_text(line.lstrip("•").strip()))

        if first == "Decor Level":
            color_headers = [normalize_text(ws.cell(r, c).value).split("\n")[0] for c in range(5, ws.max_column + 1)]
            rr = r + 1
            while rr <= ws.max_row:
                values = [normalize_text(ws.cell(rr, c).value) for c in range(1, ws.max_column + 1)]
                if values[0] == "Exterior Solid Paint":
                    break
                if not any(values):
                    rr += 1
                    continue
                if values[0] and values[0] != "Decor Level":
                    colors = {
                        color_headers[i]: values[4 + i]
                        for i in range(len(color_headers))
                        if color_headers[i]
                    }
                    interior_rows.append(
                        ColorInteriorRow(
                            decor_level=values[0],
                            seat_type=values[1],
                            seat_code=values[2],
                            seat_trim=values[3],
                            colors=colors,
                        )
                    )
                rr += 1

        if first == "Exterior Solid Paint":
            color_headers = [normalize_text(ws.cell(r, c).value).split("\n")[0] for c in range(5, ws.max_column + 1)]
            rr = r + 1
            while rr <= ws.max_row:
                values = [normalize_text(ws.cell(rr, c).value) for c in range(1, ws.max_column + 1)]
                if not any(values):
                    rr += 1
                    continue
                if FOOTNOTE_LINE_RE.match(values[0]) or values[0].startswith("•"):
                    break
                if values[0] and values[0] != "Exterior Solid Paint":
                    colors = {
                        color_headers[i]: values[4 + i]
                        for i in range(len(color_headers))
                        if color_headers[i]
                    }
                    exterior_rows.append(
                        ColorExteriorRow(
                            title=values[0],
                            color_code=values[2],
                            touch_up_paint_number=values[3],
                            colors=colors,
                        )
                    )
                rr += 1

    return ColorSheet(
        name=ws.title,
        heading_text=" | ".join(unique_preserve_order(heading_lines)),
        footnotes=footnotes,
        bullet_notes=unique_preserve_order(bullet_notes),
        interior_rows=interior_rows,
        exterior_rows=exterior_rows,
    )

def parse_spec_sheet(ws) -> List[SpecColumn]:
    section_markers = {"Specifications", "Capacities"}
    first_section_row: Optional[int] = None
    for r in range(1, min(ws.max_row, 10) + 1):
        if normalize_text(ws.cell(r, 1).value) in section_markers:
            first_section_row = r
            break
    if first_section_row is None:
        return []

    header_row = first_section_row
    nonempty_on_section_row = sum(1 for c in range(1, ws.max_column + 1) if normalize_text(ws.cell(first_section_row, c).value))
    if nonempty_on_section_row <= 1 and first_section_row > 1:
        header_row = first_section_row - 1

    header_cells = [normalize_text(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
    top_label = ""
    for r in range(1, header_row + 1):
        cell = normalize_text(ws.cell(r, 1).value)
        if cell and cell not in section_markers and "all dimensions" not in cell.lower():
            top_label = cell
            break

    columns: List[SpecColumn] = []
    for c in range(2, ws.max_column + 1):
        header = normalize_text(ws.cell(header_row, c).value)
        if not header:
            continue
        header_lines = [normalize_text(x) for x in header.split("\n") if normalize_text(x)]
        current_section = normalize_text(ws.cell(first_section_row, 1).value) if header_row == first_section_row else ""
        cells: List[SpecCell] = []
        for r in range(header_row + 1, ws.max_row + 1):
            label = normalize_text(ws.cell(r, 1).value)
            row_values = [normalize_text(ws.cell(r, cc).value) for cc in range(1, ws.max_column + 1)]
            if label in section_markers and sum(1 for x in row_values[1:] if x) == 0:
                current_section = label
                continue
            value = normalize_text(ws.cell(r, c).value)
            if label and value and value != "--":
                cells.append(SpecCell(section=current_section or "Data", label=label, value=value))
        if cells:
            columns.append(
                SpecColumn(
                    sheet_name=ws.title,
                    top_label=top_label,
                    header=header,
                    header_lines=header_lines,
                    cells=cells,
                )
            )
    return columns

def parse_engine_axles_sheet(ws) -> List[EngineAxleEntry]:
    if not ws.title.startswith("Engine Axles"):
        return []
    top_label = normalize_text(ws.cell(1, 1).value)
    legend_text = " ".join(
        normalize_text(ws.cell(r, c).value)
        for r in range(1, min(ws.max_row, 4) + 1)
        for c in range(1, ws.max_column + 1)
        if normalize_text(ws.cell(r, c).value)
    )
    footnotes: Dict[str, str] = {}
    for r in range(1, ws.max_row + 1):
        first = normalize_text(ws.cell(r, 1).value)
        if first:
            footnotes.update(parse_footnote_map(first))

    section_headers: Dict[int, str] = {}
    current_section = ""
    for c in range(3, ws.max_column + 1):
        value = normalize_text(ws.cell(3, c).value)
        if value:
            current_section = value
        section_headers[c] = current_section

    entries: List[EngineAxleEntry] = []
    current_model = ""
    for r in range(5, ws.max_row + 1):
        model = normalize_text(ws.cell(r, 1).value)
        engine = normalize_text(ws.cell(r, 2).value)
        row_values = [normalize_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if not any(row_values):
            continue
        if model and FOOTNOTE_LINE_RE.match(model):
            continue
        if model:
            current_model = model
        if not engine:
            continue
        items: List[EngineAxleItem] = []
        for c in range(3, ws.max_column + 1):
            raw_status = normalize_text(ws.cell(r, c).value)
            if not raw_status or raw_status == "--":
                continue
            code, label, notes = parse_status_value(raw_status, {}, footnotes)
            items.append(
                EngineAxleItem(
                    category=section_headers.get(c, ""),
                    name=normalize_text(ws.cell(4, c).value),
                    raw_status=raw_status,
                    status_code=code,
                    status_label=label,
                    notes=notes,
                )
            )
        if items:
            entries.append(
                EngineAxleEntry(
                    sheet_name=ws.title,
                    top_label=top_label,
                    model_code=current_model,
                    engine=engine,
                    items=items,
                )
            )
    return entries

def parse_value_and_footnote_ids(raw: str) -> Tuple[str, List[str]]:
    raw = normalize_text(raw)
    if not raw or raw == "--":
        return "", []
    m = re.match(r"^(.*\))(\d+)$", raw)
    if m:
        return normalize_text(m.group(1)), re.findall(r"\d+", m.group(2))
    m = re.match(r"^([0-9]+\.[0-9]{2})(\d+)$", raw)
    if m:
        return normalize_text(m.group(1)), re.findall(r"\d+", m.group(2))
    m = TRAILING_DIGITS_RE.match(raw)
    if m and re.search(r"[A-Za-z]", m.group(1)):
        return normalize_text(m.group(1)), re.findall(r"\d+", m.group(2))
    return raw, []

def parse_trailering_sheet(ws) -> Tuple[List[TraileringRecord], List[GCWRRecord]]:
    if not ws.title.startswith("Trailering Specs"):
        return [], []

    sheet_footnotes: Dict[str, str] = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            sheet_footnotes.update(parse_footnote_map(normalize_text(ws.cell(r, c).value)))

    rating_note = normalize_text(ws.cell(1, 1).value)
    rating_type = normalize_text(ws.cell(2, 1).value)

    engine_pairs: List[Tuple[str, int, int]] = []
    c = 2
    while c <= ws.max_column:
        engine = normalize_text(ws.cell(3, c).value)
        if engine:
            engine_pairs.append((engine, c, c + 1))
        c += 2

    trailering_records: List[TraileringRecord] = []
    current_model = ""
    gcwr_start_row = None
    for r in range(5, ws.max_row + 1):
        first = normalize_text(ws.cell(r, 1).value)
        if first.startswith("GCWR "):
            gcwr_start_row = r
            break
        if first and FOOTNOTE_LINE_RE.match(first):
            continue
        if first:
            current_model = first
        if not current_model:
            continue
        for engine, axle_col, weight_col in engine_pairs:
            raw_axle = normalize_text(ws.cell(r, axle_col).value)
            raw_weight = normalize_text(ws.cell(r, weight_col).value)
            if (not raw_axle or raw_axle == "--") and (not raw_weight or raw_weight == "--"):
                continue
            axle_ratio, axle_note_ids = parse_value_and_footnote_ids(raw_axle)
            max_weight, weight_note_ids = parse_value_and_footnote_ids(raw_weight)
            note_texts = [sheet_footnotes[nid] for nid in axle_note_ids + weight_note_ids if nid in sheet_footnotes]
            trailering_records.append(
                TraileringRecord(
                    sheet_name=ws.title,
                    rating_type=rating_type,
                    note_text=rating_note,
                    model_code=current_model,
                    engine=engine,
                    axle_ratio=axle_ratio or raw_axle,
                    max_trailer_weight=max_weight or raw_weight,
                    footnotes=unique_preserve_order(note_texts),
                )
            )

    gcwr_records: List[GCWRRecord] = []
    if gcwr_start_row is not None and gcwr_start_row + 3 <= ws.max_row:
        table_title = normalize_text(ws.cell(gcwr_start_row, 1).value)
        header_values = [normalize_text(ws.cell(gcwr_start_row + 2, c).value) for c in range(2, ws.max_column + 1)]
        for r in range(gcwr_start_row + 3, ws.max_row + 1):
            engine = normalize_text(ws.cell(r, 1).value)
            if not engine or FOOTNOTE_LINE_RE.match(engine):
                continue
            for idx, gcwr in enumerate(header_values, start=2):
                if not gcwr:
                    continue
                raw_axle = normalize_text(ws.cell(r, idx).value)
                if not raw_axle or raw_axle == "--":
                    continue
                axle_ratio, note_ids = parse_value_and_footnote_ids(raw_axle)
                gcwr_records.append(
                    GCWRRecord(
                        sheet_name=ws.title,
                        table_title=table_title,
                        engine=engine,
                        gcwr=gcwr,
                        axle_ratio=axle_ratio or raw_axle,
                        footnotes=unique_preserve_order(sheet_footnotes[nid] for nid in note_ids if nid in sheet_footnotes),
                    )
                )

    return trailering_records, gcwr_records

def parse_glossary_sheet(ws) -> OrderedDict[str, str]:
    glossary: OrderedDict[str, str] = OrderedDict()
    first = normalize_text(ws.cell(1, 1).value)
    second = normalize_text(ws.cell(1, 2).value)
    if first != "Option Code" or second != "Description":
        return glossary
    for r in range(2, ws.max_row + 1):
        code = normalize_text(ws.cell(r, 1).value)
        description = normalize_text(ws.cell(r, 2).value)
        if code and description:
            glossary[code] = description
    return glossary

def parse_workbook(path: Path) -> WorkbookData:
    wb = openpyxl.load_workbook(path, data_only=True)
    year, make, model, vehicle_name = parse_filename_metadata(path)

    trim_defs: List[TrimDef] = []
    matrix_sheets: List[MatrixSheet] = []
    color_sheets: List[ColorSheet] = []
    spec_columns: List[SpecColumn] = []
    engine_axle_entries: List[EngineAxleEntry] = []
    trailering_records: List[TraileringRecord] = []
    gcwr_records: List[GCWRRecord] = []
    glossary: OrderedDict[str, str] = OrderedDict()

    for name in wb.sheetnames:
        ws = wb[name]
        matrix = parse_matrix_sheet(ws, trim_defs or None)
        if matrix:
            if not trim_defs:
                trim_defs = matrix.trim_defs
            matrix_sheets.append(matrix)
        if name.startswith("Colour and Trim"):
            color_sheets.append(parse_color_sheet(ws))
        if name.startswith("Dimensions") or name.startswith("Specs") or name in {"Dimensions", "Specs"}:
            spec_columns.extend(parse_spec_sheet(ws))
        if name.startswith("Engine Axles"):
            engine_axle_entries.extend(parse_engine_axles_sheet(ws))
        if name.startswith("Trailering Specs"):
            records, gcwrs = parse_trailering_sheet(ws)
            trailering_records.extend(records)
            gcwr_records.extend(gcwrs)
        if name == "All":
            glossary.update(parse_glossary_sheet(ws))

    return WorkbookData(
        path=path,
        year=year,
        make=make,
        model=model,
        vehicle_name=vehicle_name,
        trim_defs=trim_defs,
        matrix_sheets=matrix_sheets,
        color_sheets=color_sheets,
        spec_columns=spec_columns,
        engine_axle_entries=engine_axle_entries,
        trailering_records=trailering_records,
        gcwr_records=gcwr_records,
        glossary=glossary,
        sheet_names=wb.sheetnames,
    )

def referenced_codes_for_text(text: str, glossary: Dict[str, str]) -> List[Tuple[str, str]]:
    codes = []
    for code in CODE_IN_PARENS_RE.findall(text):
        if code in glossary:
            codes.append((code, glossary[code]))
    return list(OrderedDict(((code, desc), None) for code, desc in codes).keys())
