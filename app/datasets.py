from __future__ import annotations

from dataclasses import asdict, dataclass
import json
import os
import re
from typing import Any

import pandas as pd

from .config import (
    get_dataset_meta,
    get_dataset_paths,
    load_admin_config,
    save_dataset_meta,
    source_signature,
    ensure_dirs,
)
from .utils import (
    clean_text,
    compact_text,
    extract_parenthetical_buildings,
    normalize_range_symbols,
    parse_block_unit_ranges,
    parse_parcel_ranges,
    tokenize_korean_terms,
)


@dataclass
class AdminDistrictRecord:
    dataset_version: int
    row_number: int
    admin_area: str
    tong_ri: str
    ban: str
    jurisdiction: str
    legal_area: str
    building_name: str
    building_names: list[str]
    block_numbers: list[str]
    unit_ranges: list[dict]
    token_text: str
    parcel_ranges: list[dict]


@dataclass
class ReportTemplateRecord:
    sheet_name: str
    row_number: int
    school_name: str
    admin_area: str
    tong_ri: str
    extra_area: str
    note: str
    category: str
    tong_key: str
    extra_tokens: list[str]


@dataclass
class SchoolZoneRecord:
    row_number: int
    school_name: str
    admin_area: str
    zone_text: str
    note: str
    change_note: str
    tong_key: str
    zone_tokens: list[str]


def _is_stale(target: str, sources: list[str]) -> bool:
    if not os.path.exists(target):
        return True
    target_mtime = os.path.getmtime(target)
    return any(os.path.getmtime(path) > target_mtime for path in sources if os.path.exists(path))


def _write_json(path: str, records: list[dict[str, Any]]) -> None:
    ensure_dirs()
    with open(path, "w", encoding="utf-8") as file:
        json.dump(records, file, ensure_ascii=False, indent=2)


def _read_json(path: str) -> list[dict[str, Any]]:
    with open(path, "r", encoding="utf-8") as file:
        return json.load(file)


def _source_exists(path: str) -> bool:
    return bool(path) and os.path.exists(path)


def _load_or_build_dataset(
    dataset_path: str,
    source_paths: list[str],
    build_fn,
    force_rebuild: bool = False,
):
    source_ready = all(_source_exists(path) for path in source_paths)
    dataset_exists = os.path.exists(dataset_path)

    if force_rebuild and source_ready:
        return build_fn()
    if source_ready and _is_stale(dataset_path, source_paths):
        return build_fn()
    if dataset_exists:
        return _read_json(dataset_path)
    if source_ready:
        return build_fn()
    missing = ", ".join(path for path in source_paths if not _source_exists(path))
    raise FileNotFoundError(f"필수 원본 파일이 없고 사전 생성 데이터셋도 없습니다: {missing}")


def build_admin_dataset() -> list[dict[str, Any]]:
    config = load_admin_config()
    dataset_paths = get_dataset_paths()
    df = pd.read_excel(config.data_sources.admin_source, sheet_name="sheet1", header=None)
    dataset_version = 3
    current_admin_area = ""
    current_tong_ri = ""
    current_ban = ""
    current_legal_area = ""
    records: list[AdminDistrictRecord] = []

    for index, row in df.iterrows():
        col0 = clean_text(row.iloc[0] if len(row) > 0 else "")
        col1 = clean_text(row.iloc[1] if len(row) > 1 else "")
        col2 = clean_text(row.iloc[2] if len(row) > 2 else "")
        col3 = clean_text(row.iloc[3] if len(row) > 3 else "")

        if index < 4:
            continue

        if col0 and any(key in col0 for key in ["동", "읍", "면"]):
            if not any(key in col0 for key in ["총계", "소계"]):
                current_admin_area = col0
                current_tong_ri = ""
                current_ban = ""
                current_legal_area = ""
            continue

        if col1 and ("리" in col1 or "통" in col1):
            current_tong_ri = clean_text(col1).replace("제", "")
            current_ban = ""
            current_legal_area = ""

        if not col2 and not col3:
            continue
        if "총계" in col1 or "소계" in col1:
            continue
        if not current_admin_area or not current_tong_ri or not col3:
            continue

        jurisdiction = normalize_range_symbols(col3)
        legal_match = re.search(r"([가-힣]+(?:동|리))", jurisdiction)
        if col2:
            current_ban = clean_text(col2).replace("제", "")
        ban = current_ban
        if legal_match:
            current_legal_area = legal_match.group(1)
        legal_area = current_legal_area

        building_match = re.search(
            r"([가-힣A-Za-z0-9]+(?:아파트|마을\d*단지|단지|빌라|타운|파크|캐슬|하우스|하임|팰리스|레지던스|휴먼빌|푸르지오|자이|e편한세상))",
            jurisdiction,
        )
        building_name = building_match.group(1) if building_match else ""
        building_names = extract_parenthetical_buildings(jurisdiction)
        if building_name and building_name not in building_names:
            building_names.append(building_name)
        block_numbers, unit_ranges = parse_block_unit_ranges(jurisdiction)
        parcel_ranges = parse_parcel_ranges(jurisdiction, legal_area=legal_area)

        token_text = compact_text(" ".join(filter(None, [current_admin_area, current_tong_ri, col2, jurisdiction])))

        records.append(
            AdminDistrictRecord(
                dataset_version=dataset_version,
                row_number=index + 1,
                admin_area=current_admin_area,
                tong_ri=current_tong_ri,
                ban=ban,
                jurisdiction=jurisdiction,
                legal_area=legal_area,
                building_name=building_name,
                building_names=building_names,
                block_numbers=block_numbers,
                unit_ranges=unit_ranges,
                token_text=token_text,
                parcel_ranges=parcel_ranges,
            )
        )

    serialized = [asdict(record) for record in records]
    _write_json(dataset_paths["admin"], serialized)
    return serialized


def build_report_template_dataset() -> list[dict[str, Any]]:
    config = load_admin_config()
    dataset_paths = get_dataset_paths()
    xls = pd.ExcelFile(config.data_sources.report_template_source)
    records: list[ReportTemplateRecord] = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(config.data_sources.report_template_source, sheet_name=sheet_name, header=None)
        current_school = ""
        for index, row in df.iterrows():
            row_number = index + 1
            school_name = clean_text(row.iloc[0] if len(row) > 0 else "")
            admin_area = clean_text(row.iloc[1] if len(row) > 1 else "")
            tong_ri = clean_text(row.iloc[2] if len(row) > 2 else "")
            extra_area = clean_text(row.iloc[3] if len(row) > 3 else "")
            note = clean_text(row.iloc[11] if sheet_name == "재학생 현황(관내)" and len(row) > 11 else "")
            if sheet_name == "재학생 현황(관외)":
                note = clean_text(row.iloc[12] if len(row) > 12 else "")

            if school_name:
                current_school = school_name

            if row_number <= 5 or not current_school:
                continue

            if admin_area in {"합계", "기타(통학구역외)", "제한적 공동통학구역"}:
                category = "summary" if admin_area == "합계" else "special"
            else:
                category = "normal"

            if not any([admin_area, tong_ri, extra_area]) and category == "normal":
                continue

            tokens = tokenize_korean_terms(extra_area)
            if not tokens and "(" in tong_ri:
                inside = re.search(r"\(([^)]+)\)", tong_ri)
                if inside:
                    tokens = tokenize_korean_terms(inside.group(1))

            records.append(
                ReportTemplateRecord(
                    sheet_name=sheet_name,
                    row_number=row_number,
                    school_name=current_school,
                    admin_area=admin_area,
                    tong_ri=tong_ri,
                    extra_area=extra_area,
                    note=note,
                    category=category,
                    tong_key=compact_text(tong_ri),
                    extra_tokens=tokens,
                )
            )

    serialized = [asdict(record) for record in records]
    _write_json(dataset_paths["report_template"], serialized)
    return serialized


def build_school_zone_dataset() -> list[dict[str, Any]]:
    config = load_admin_config()
    dataset_paths = get_dataset_paths()
    df = pd.read_excel(config.data_sources.school_zone_source, sheet_name=0, header=None)
    records: list[SchoolZoneRecord] = []
    for index, row in df.iterrows():
        if index < 7:
            continue
        school_name = clean_text(row.iloc[1] if len(row) > 1 else "")
        admin_area = clean_text(row.iloc[2] if len(row) > 2 else "")
        zone_text = clean_text(row.iloc[3] if len(row) > 3 else "")
        note = clean_text(row.iloc[4] if len(row) > 4 else "")
        change_note = clean_text(row.iloc[5] if len(row) > 5 else "")
        if not school_name or not admin_area or not zone_text:
            continue
        records.append(
            SchoolZoneRecord(
                row_number=index + 1,
                school_name=school_name,
                admin_area=admin_area,
                zone_text=normalize_range_symbols(zone_text),
                note=note,
                change_note=change_note,
                tong_key=compact_text(zone_text),
                zone_tokens=tokenize_korean_terms(zone_text),
            )
        )
    serialized = [asdict(record) for record in records]
    _write_json(dataset_paths["school_zones"], serialized)
    return serialized


def ensure_datasets() -> dict[str, list[dict[str, Any]]]:
    ensure_dirs()
    config = load_admin_config()
    dataset_paths = get_dataset_paths()
    sources = source_signature(config)
    meta = get_dataset_meta()
    config_changed = meta.get("source_signature") != sources

    admin = _load_or_build_dataset(
        dataset_path=dataset_paths["admin"],
        source_paths=[config.data_sources.admin_source],
        build_fn=build_admin_dataset,
        force_rebuild=config_changed,
    )
    if admin and admin[0].get("dataset_version") != 3 and _source_exists(config.data_sources.admin_source):
        admin = build_admin_dataset()

    report = _load_or_build_dataset(
        dataset_path=dataset_paths["report_template"],
        source_paths=[config.data_sources.report_template_source],
        build_fn=build_report_template_dataset,
        force_rebuild=config_changed,
    )

    zones = _load_or_build_dataset(
        dataset_path=dataset_paths["school_zones"],
        source_paths=[config.data_sources.school_zone_source],
        build_fn=build_school_zone_dataset,
        force_rebuild=config_changed,
    )
    if zones and ("tong_key" not in zones[0] or "zone_tokens" not in zones[0]) and _source_exists(config.data_sources.school_zone_source):
        zones = build_school_zone_dataset()

    save_dataset_meta({"source_signature": sources})

    return {
        "admin": admin,
        "report_template": report,
        "school_zones": zones,
    }
