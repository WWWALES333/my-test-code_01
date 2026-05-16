from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from .schema import OWNER_TYPE_PERSON, stable_hash

XLSX_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def discover_roster_file(explicit_path: Path | None = None) -> Path | None:
    """返回首个可用花名册文件。"""
    if explicit_path and explicit_path.exists():
        return explicit_path
    base = Path("data/input/v1.5/roster")
    if not base.exists():
        return None
    files = sorted(base.glob("*.xlsx"))
    return files[0] if files else None


def load_sales_roster(path: Path) -> List[Dict[str, object]]:
    """读取花名册 Excel，产出当前在岗销售主数据。"""
    rows = _read_xlsx_rows(path)
    if not rows:
        return []
    header = [str(item).strip() for item in rows[0]]
    index = {name: idx for idx, name in enumerate(header)}
    required = ("花名", "汤名", "组织全称", "部门", "职务")
    missing = [name for name in required if name not in index]
    if missing:
        raise ValueError(f"花名册缺少字段: {', '.join(missing)}")

    source_date = _infer_source_date(path)
    roster: List[Dict[str, object]] = []
    seen_ids: set[str] = set()
    for raw in rows[1:]:
        flower_name = _cell(raw, index, "花名")
        if not flower_name:
            continue
        team_name = _cell(raw, index, "汤名")
        org_full_name = _cell(raw, index, "组织全称")
        department_name = _cell(raw, index, "部门")
        job_title = _cell(raw, index, "职务")
        battle_zone = _extract_battle_zone(team_name, org_full_name, department_name)
        region_name = _extract_region_name(org_full_name, department_name)
        salesperson_id = stable_hash("roster", flower_name, battle_zone or region_name or team_name, length=12)
        if salesperson_id in seen_ids:
            continue
        seen_ids.add(salesperson_id)
        roster.append(
            {
                "salesperson_id": salesperson_id,
                "display_name": flower_name,
                "salesperson_name": flower_name,
                "flower_name": flower_name,
                "team_name": team_name,
                "org_full_name": org_full_name,
                "department_name": department_name,
                "job_title": job_title,
                "battle_zone_name": battle_zone,
                "region_name": region_name,
                "employment_status": "active",
                "roster_source_date": source_date,
                "owner_type": OWNER_TYPE_PERSON,
                "owner_hint": flower_name,
                "aliases": [flower_name],
            }
        )
    return sorted(roster, key=lambda item: (str(item.get("battle_zone_name", "")), str(item.get("region_name", "")), str(item.get("display_name", ""))))


def build_roster_lookup(roster_rows: Iterable[Dict[str, object]]) -> Dict[str, Dict[str, object]]:
    """按花名和别名构建销售主数据查询表。"""
    lookup: Dict[str, Dict[str, object]] = {}
    for row in roster_rows:
        for token in (str(row.get("flower_name", "")), str(row.get("display_name", "")), *[str(item) for item in row.get("aliases", [])]):
            key = token.strip()
            if key:
                lookup[key] = row
    return lookup


def _read_xlsx_rows(path: Path) -> List[List[str]]:
    rows: List[List[str]] = []
    with ZipFile(path) as archive:
        shared_strings = _read_shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.findall(".//a:sheets/a:sheet", XLSX_NS)
        if not sheets:
            return rows
        first_sheet_id = sheets[0].attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target = ""
        for rel in rels:
            if rel.attrib.get("Id") == first_sheet_id:
                target = rel.attrib.get("Target", "")
                break
        if not target:
            target = "worksheets/sheet1.xml"
        sheet_path = f"xl/{target.lstrip('/')}"
        sheet = ET.fromstring(archive.read(sheet_path))
        for row in sheet.findall(".//a:sheetData/a:row", XLSX_NS):
            values: List[str] = []
            for cell in row.findall("a:c", XLSX_NS):
                cell_type = cell.attrib.get("t")
                value = cell.find("a:v", XLSX_NS)
                if value is None or value.text is None:
                    values.append("")
                    continue
                if cell_type == "s":
                    values.append(shared_strings[int(value.text)])
                else:
                    values.append(value.text)
            rows.append(values)
    return rows


def _read_shared_strings(archive: ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: List[str] = []
    for item in root.findall("a:si", XLSX_NS):
        strings.append("".join(node.text or "" for node in item.iterfind(".//a:t", XLSX_NS)))
    return strings


def _infer_source_date(path: Path) -> str:
    matched = re.search(r"(20\d{2})(\d{2})(\d{2})", path.stem)
    if not matched:
        return ""
    return f"{matched.group(1)}-{matched.group(2)}-{matched.group(3)}"


def _extract_battle_zone(team_name: str, org_full_name: str, department_name: str) -> str:
    for text in (team_name, org_full_name, department_name):
        matched = re.search(r"(将军汤[^/]{0,16}战区(?:（[^）]+）)?)", text)
        if matched:
            return matched.group(1).strip()
        matched = re.search(r"(销售[^/]{0,16}战区(?:（[^）]+）)?)", text)
        if matched:
            return matched.group(1).strip()
    return ""


def _extract_region_name(org_full_name: str, department_name: str) -> str:
    for text in (department_name, org_full_name):
        matched = re.search(r"([\u4e00-\u9fa5]{2,12}区域)$", text)
        if matched:
            return matched.group(1).strip()
    return department_name.strip()


def _cell(row: List[str], index: Dict[str, int], field_name: str) -> str:
    idx = index.get(field_name, -1)
    if idx < 0 or idx >= len(row):
        return ""
    return str(row[idx]).strip()
