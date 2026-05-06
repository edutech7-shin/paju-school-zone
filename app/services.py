from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
import os
import re
import threading
import time
from typing import Any
from uuid import uuid4

import pandas as pd
from openpyxl import load_workbook

from addressConverter import get_jibeon_address
from mapping_final import (
    extract_address_components,
    extract_building_name,
    parse_district_ranges,
    map_address_to_school_area,
    manual_mapping,
    standardize_address,
)

from .config import load_admin_config
from .datasets import ensure_datasets
from .utils import any_contains, clean_text, compact_text, parcel_in_ranges, parse_jibun_address, tokenize_korean_terms
from .utils import parse_address_unit_info, unit_in_ranges


GRADE_COLUMNS = [f"{grade}학년" for grade in range(1, 7)]


@dataclass
class AppContext:
    admin_records: list[dict[str, Any]]
    report_template_records: list[dict[str, Any]]
    school_zone_records: list[dict[str, Any]]
    fallback_district_ranges: pd.DataFrame


_CONTEXT_CACHE: dict[tuple[str, str, str, str], AppContext] = {}
_CLEANUP_THREAD_STARTED = False


def get_context() -> AppContext:
    config = load_admin_config()
    key = (
        config.data_sources.admin_source,
        config.data_sources.report_template_source,
        config.data_sources.school_zone_source,
        config.data_sources.legacy_district_source,
    )
    if key in _CONTEXT_CACHE:
        return _CONTEXT_CACHE[key]
    datasets = ensure_datasets()
    fallback_district_ranges = _load_fallback_district_ranges(config.data_sources.legacy_district_source)
    context = AppContext(
        admin_records=datasets["admin"],
        report_template_records=datasets["report_template"],
        school_zone_records=datasets["school_zones"],
        fallback_district_ranges=fallback_district_ranges,
    )
    _CONTEXT_CACHE.clear()
    _CONTEXT_CACHE[key] = context
    return context


def _load_fallback_district_ranges(source_path: str) -> pd.DataFrame:
    """
    Vercel 배포 환경에서는 로컬 절대경로 원본 엑셀이 없을 수 있으므로,
    파일이 없거나 읽기에 실패해도 빈 범위 테이블로 안전하게 처리합니다.
    """
    empty_columns = ["동이름", "관할동", "관할통", "시작본번", "시작부번", "끝본번", "끝부번"]
    if not source_path or not os.path.exists(source_path):
        return pd.DataFrame(columns=empty_columns)
    try:
        district_df = pd.read_excel(source_path)
        parsed = parse_district_ranges(district_df)
        if parsed is None or parsed.empty:
            return pd.DataFrame(columns=empty_columns)
        return parsed
    except Exception:
        return pd.DataFrame(columns=empty_columns)


def clear_context_cache() -> None:
    _CONTEXT_CACHE.clear()


def _safe_int(value, default=0) -> int:
    try:
        if pd.isna(value):
            return default
    except TypeError:
        pass
    text = str(value).strip()
    if not text:
        return default
    match = re.search(r"\d+", text)
    return int(match.group(0)) if match else default


def _format_base_date(base_date: str) -> str:
    text = clean_text(base_date)
    if not text:
        return ""
    try:
        parsed = datetime.strptime(text, "%Y-%m-%d")
        return f"{parsed.year}. {parsed.month}. {parsed.day}."
    except ValueError:
        return text


def _find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized = {compact_text(col): col for col in df.columns}
    for candidate in candidates:
        key = compact_text(candidate)
        if key in normalized:
            return normalized[key]
    return None


def _extract_student_rows(df: pd.DataFrame, fallback_grade: int | None = None) -> list[dict[str, Any]]:
    grade_col = _find_column(df, ["학년"])
    name_col = _find_column(df, ["성명", "이름", "학생명"])
    zip_col = _find_column(df, ["우편번호", "우편 번호", "zip", "zipcode"])
    address_col = _find_column(df, ["주소", "도로명주소", "주소지"])
    jibun_col = _find_column(df, ["지번주소"])

    if address_col is None and jibun_col is None:
        raise ValueError("학생 파일에 '주소' 또는 '지번주소' 컬럼이 필요합니다.")

    rows = []
    for _, raw in df.iterrows():
        grade = _safe_int(raw.get(grade_col) if grade_col else fallback_grade, fallback_grade or 0)
        name = clean_text(raw.get(name_col) if name_col else "")
        postal_code = clean_text(raw.get(zip_col) if zip_col else "")
        road_address = clean_text(raw.get(address_col) if address_col else "")
        jibun_address = clean_text(raw.get(jibun_col) if jibun_col else "")

        if not road_address and not jibun_address:
            continue

        rows.append(
            {
                "grade": grade,
                "name": name,
                "postal_code": postal_code,
                "road_address": road_address,
                "source_jibun_address": jibun_address,
            }
        )
    return rows


def load_student_files(files: list[tuple[str, bytes]]) -> list[dict[str, Any]]:
    all_rows: list[dict[str, Any]] = []
    for filename, payload in files:
        xls = pd.ExcelFile(BytesIO(payload))
        fallback_grade = _safe_int(filename, 0) or None
        file_rows = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(BytesIO(payload), sheet_name=sheet_name)
            file_rows.extend(_extract_student_rows(df, fallback_grade=fallback_grade))
        all_rows.extend(file_rows)
    if not all_rows:
        raise ValueError("학생 주소 데이터를 읽지 못했습니다.")
    return all_rows


def _extract_tong_number(text: str) -> str:
    match = re.search(r"(\d+(?:~\d+)?)통", clean_text(text))
    return match.group(1) if match else ""


def _extract_tong_numbers(text: str) -> set[int]:
    numbers: set[int] = set()
    for start_text, end_text in re.findall(r"(\d+)(?:~(\d+))?통", clean_text(text)):
        start = int(start_text)
        end = int(end_text or start_text)
        if start > end:
            start, end = end, start
        numbers.update(range(start, end + 1))
    return numbers


def _tong_matches(student_tong: str, candidate_tong: str) -> bool:
    student_numbers = _extract_tong_numbers(student_tong)
    candidate_numbers = _extract_tong_numbers(candidate_tong)
    if not student_numbers or not candidate_numbers:
        return False
    return bool(student_numbers & candidate_numbers)


def _parse_school_names(text: str) -> set[str]:
    return {token.strip() for token in re.split(r"[,/\n]+", clean_text(text)) if token.strip().endswith("초")}


def _is_joint_zone_pair(requested_row: dict[str, Any], actual_row: dict[str, Any], requested_school_name: str, actual_school_name: str) -> bool:
    if not requested_row.get("row_number") or not actual_row.get("row_number"):
        return False
    if requested_row.get("admin_area") != actual_row.get("admin_area"):
        return False
    if not _tong_matches(requested_row.get("tong_ri", ""), actual_row.get("tong_ri", "")):
        return False
    if compact_text(requested_row.get("extra_area", "")) != compact_text(actual_row.get("extra_area", "")):
        return False
    requested_note_schools = _parse_school_names(requested_row.get("note", ""))
    actual_note_schools = _parse_school_names(actual_row.get("note", ""))
    return actual_school_name in requested_note_schools or requested_school_name in actual_note_schools


def _score_school_zone_record(
    record: dict[str, Any],
    admin_area: str,
    tong_ri: str,
    building_name: str,
    address_text: str,
) -> int:
    score = 0
    address_compact = compact_text(address_text)
    building_tokens = tokenize_korean_terms(building_name)
    tong_number = _extract_tong_number(tong_ri)
    zone_text = record["zone_text"]

    if admin_area and record["admin_area"] == admin_area:
        score += 60

    if tong_number:
        if _tong_matches(tong_ri, zone_text):
            score += 70

    if tong_ri and compact_text(tong_ri) in record["tong_key"]:
        score += 35

    zone_compact = compact_text(zone_text)
    for token in building_tokens:
        if token and token in zone_compact:
            score += 30

    for token in record.get("zone_tokens", []):
        if token and token in address_compact:
            score += 12

    if zone_compact and zone_compact in address_compact:
        score += 80

    return score


def choose_school_and_zone(
    context: AppContext,
    admin_area: str,
    tong_ri: str,
    building_name: str,
    address_text: str,
    school_name: str = "",
) -> dict[str, Any]:
    candidates = context.school_zone_records
    if school_name:
        candidates = [record for record in candidates if record["school_name"] == school_name]

    scored = []
    for record in candidates:
        score = _score_school_zone_record(record, admin_area, tong_ri, building_name, address_text)
        if score > 0:
            scored.append((score, record))

    fallback_status = ""
    if not scored and school_name:
        fallback_status = "학교명불일치후재탐색"
        for record in context.school_zone_records:
            score = _score_school_zone_record(record, admin_area, tong_ri, building_name, address_text)
            if score > 0:
                scored.append((score, record))

    if not scored:
        return {"school_name": school_name, "zone_status": "학교미분류"}

    scored.sort(key=lambda item: item[0], reverse=True)
    best_score, best = scored[0]
    if fallback_status:
        status = fallback_status
    else:
        status = "학교매칭" if best_score >= 100 else "학교후보"
    return {**best, "zone_status": status, "zone_score": best_score}


def _admin_lookup_from_dataset(address_text: str, building_name: str, context: AppContext) -> tuple[str, str, str, dict[str, Any] | None]:
    normalized = compact_text(address_text)
    building_tokens = tokenize_korean_terms(building_name)
    address_tokens = tokenize_korean_terms(address_text)
    parsed_parcel = parse_jibun_address(address_text)
    unit_info = parse_address_unit_info(address_text)
    address_block_numbers = unit_info["block_numbers"]
    address_unit_numbers = unit_info["unit_numbers"]

    if parsed_parcel["main_number"] is not None:
        parcel_candidates: list[dict[str, Any]] = []
        for record in context.admin_records:
            if not record["legal_area"]:
                continue
            if record["legal_area"] != parsed_parcel["legal_area"]:
                continue
            if parcel_in_ranges(parsed_parcel, record.get("parcel_ranges", [])):
                parcel_candidates.append(record)
        if parcel_candidates:
            if len(parcel_candidates) == 1:
                record = parcel_candidates[0]
                return record["admin_area"], record["tong_ri"], record["ban"], record

            def score_candidate(record: dict[str, Any]) -> tuple[int, int]:
                candidate_tokens = tokenize_korean_terms(" ".join(record.get("building_names", [])))
                score = 0
                for token in building_tokens + address_tokens:
                    if token and any(token in candidate for candidate in candidate_tokens):
                        score += 100
                if building_name and any_contains(" ".join(record.get("building_names", [])), [building_name]):
                    score += 140
                if record.get("building_name") and compact_text(record["building_name"]) in normalized:
                    score += 80
                record_blocks = [clean_text(value) for value in record.get("block_numbers", [])]
                if address_block_numbers and record_blocks:
                    if any(block in record_blocks for block in address_block_numbers):
                        score += 160
                    else:
                        score -= 80
                if address_unit_numbers and record.get("unit_ranges"):
                    unit_score = 0
                    for unit_number in address_unit_numbers:
                        if unit_in_ranges(unit_number, record.get("unit_ranges", []), address_block_numbers[0] if address_block_numbers else ""):
                            unit_score = max(unit_score, 120)
                    score += unit_score
                if record["jurisdiction"] and compact_text(record["jurisdiction"]) in normalized:
                    score += 30
                return score, -record["row_number"]

            best = max(parcel_candidates, key=score_candidate)
            record = best
            return record["admin_area"], record["tong_ri"], record["ban"], record

    if building_tokens or address_tokens:
        for record in context.admin_records:
            candidate_tokens = record.get("building_names", [])
            if candidate_tokens and any_contains(" ".join(candidate_tokens), building_tokens + address_tokens):
                if record["legal_area"] and record["legal_area"] in clean_text(address_text):
                    return record["admin_area"], record["tong_ri"], record["ban"], record

    for record in context.admin_records:
        if record["legal_area"] and record["legal_area"] not in clean_text(address_text):
            continue
        if record["jurisdiction"] and compact_text(record["jurisdiction"]) in normalized:
            return record["admin_area"], record["tong_ri"], record["ban"], record

    dong, main, sub = extract_address_components(clean_text(address_text))
    parcel_key = f"{dong} {main}" + (f"-{sub}" if sub else "") if dong and main else ""
    if parcel_key:
        for record in context.admin_records:
            if parcel_key and compact_text(parcel_key) in compact_text(record["jurisdiction"]):
                return record["admin_area"], record["tong_ri"], record["ban"], record

    return "", "", "", None


def _fallback_admin_lookup(jibun_address: str, context: AppContext) -> tuple[str, str]:
    dong_name, main_number, sub_number = extract_address_components(jibun_address)
    row = pd.Series(
        {
            "지번주소": jibun_address,
            "std_address": standardize_address(jibun_address),
            "building_name": extract_building_name(jibun_address),
            "dong_name": dong_name,
            "main_number": main_number,
            "sub_number": sub_number,
        }
    )
    mapped_dong, mapped_tong = map_address_to_school_area(
        row,
        manual_mapping,
        {},
        {},
        context.fallback_district_ranges,
    )
    return mapped_dong, mapped_tong


def classify_student(student: dict[str, Any], context: AppContext, school_name: str = "") -> dict[str, Any]:
    road_address = student["road_address"]
    postal_code = student.get("postal_code", "")
    requested_school_name = school_name
    source_jibun = student["source_jibun_address"]
    base_address = source_jibun or road_address
    is_jibun = bool(source_jibun) or bool(re.search(r"[가-힣]+동\s*\d+", base_address))

    jibun_address = source_jibun
    if not jibun_address and road_address:
        jibun_address = get_jibeon_address(road_address, expected_zip=postal_code)

    match_text = jibun_address or road_address
    building_name = extract_building_name(match_text)
    admin_area, tong_ri, ban, admin_record = _admin_lookup_from_dataset(match_text, building_name, context)
    if (not admin_area or not tong_ri) and jibun_address:
        fallback_dong, fallback_tong = _fallback_admin_lookup(jibun_address, context)
        admin_area = admin_area or fallback_dong
        tong_ri = tong_ri or fallback_tong

    school_zone_match = choose_school_and_zone(
        context=context,
        admin_area=admin_area,
        tong_ri=tong_ri,
        building_name=building_name,
        address_text=" ".join(filter(None, [road_address, jibun_address])),
        school_name="",
    )
    resolved_school_name = school_zone_match.get("school_name", "") or requested_school_name

    report_match = match_report_row(
        context=context,
        admin_area=admin_area,
        tong_ri=tong_ri,
        building_name=building_name,
        address_text=" ".join(filter(None, [road_address, jibun_address])),
        school_name="",
        admin_building_names=admin_record.get("building_names", []) if admin_record else [],
    )
    actual_school_name = report_match.get("school_name", "") or resolved_school_name
    if actual_school_name:
        refined_report_match = match_report_row(
            context=context,
            admin_area=admin_area,
            tong_ri=tong_ri,
            building_name=building_name,
            address_text=" ".join(filter(None, [road_address, jibun_address])),
            school_name=actual_school_name,
            admin_building_names=admin_record.get("building_names", []) if admin_record else [],
        )
        if refined_report_match.get("row_number"):
            report_match = refined_report_match
            actual_school_name = report_match.get("school_name", "") or actual_school_name
    requested_report_match = {}
    if requested_school_name:
        requested_report_match = match_report_row(
            context=context,
            admin_area=admin_area,
            tong_ri=tong_ri,
            building_name=building_name,
            address_text=" ".join(filter(None, [road_address, jibun_address])),
            school_name=requested_school_name,
            admin_building_names=admin_record.get("building_names", []) if admin_record else [],
        )
        if (
            requested_school_name
            and actual_school_name
            and requested_school_name != actual_school_name
            and _is_joint_zone_pair(requested_report_match, report_match, requested_school_name, actual_school_name)
        ):
            report_match = requested_report_match
            actual_school_name = requested_school_name
    school_alignment_status = "일치"
    zone_match_status = school_zone_match.get("zone_status", "학교미분류")
    match_status = report_match.get("status", "미분류")

    if requested_school_name and actual_school_name and requested_school_name != actual_school_name:
        school_alignment_status = "학구불일치"
        zone_match_status = "학구불일치"
        match_status = "학구불일치"

    return {
        "grade": student["grade"],
        "name": student["name"],
        "input_type": "jibun" if is_jibun else "road",
        "requested_school_name": requested_school_name,
        "postal_code": postal_code,
        "road_address": road_address,
        "jibun_address": jibun_address,
        "building_name": building_name,
        "admin_area": admin_area,
        "tong_ri": tong_ri,
        "ban": ban,
        "admin_building_names": admin_record.get("building_names", []) if admin_record else [],
        "school_alignment_status": school_alignment_status,
        "school_name": actual_school_name,
        "sheet_name": report_match.get("sheet_name", "재학생 현황(관내)"),
        "report_row_number": report_match.get("row_number"),
        "report_admin_area": report_match.get("admin_area", ""),
        "report_tong_ri": report_match.get("tong_ri", ""),
        "report_extra_area": report_match.get("extra_area", ""),
        "zone_match_status": zone_match_status,
        "zone_match_score": school_zone_match.get("zone_score", 0),
        "match_status": match_status,
    }


def match_report_row(
    context: AppContext,
    admin_area: str,
    tong_ri: str,
    building_name: str,
    address_text: str,
    school_name: str = "",
    admin_building_names: list[str] | None = None,
) -> dict[str, Any]:
    candidate_rows = [
        record
        for record in context.report_template_records
        if record["sheet_name"] == "재학생 현황(관내)" and record["category"] == "normal"
    ]

    if school_name:
        candidate_rows = [record for record in candidate_rows if record["school_name"] == school_name]

    if admin_area:
        candidate_rows = [record for record in candidate_rows if record["admin_area"] == admin_area]

    tong_number = _extract_tong_number(tong_ri)
    if tong_number:
        filtered = []
        for record in candidate_rows:
            if _tong_matches(tong_ri, record["tong_ri"]):
                filtered.append(record)
        if filtered:
            candidate_rows = filtered

    if not candidate_rows:
        return {"status": "미분류"}

    if len(candidate_rows) > 1:
        exclude_rows = [row for row in candidate_rows if "제외" in row["extra_area"]]
        explicit_rows = [row for row in candidate_rows if row["extra_area"] and "제외" not in row["extra_area"]]
        combined_tokens = tokenize_korean_terms(building_name) + tokenize_korean_terms(" ".join(admin_building_names or []))
        explicit_match = []
        for row in explicit_rows:
            row_tokens = tokenize_korean_terms(row["extra_area"])
            if any(token and any(token in candidate for candidate in row_tokens) for token in combined_tokens):
                explicit_match.append(row)
        if exclude_rows and explicit_rows and not explicit_match:
            candidate_rows = exclude_rows
        elif explicit_match:
            candidate_rows = explicit_match

    if len(candidate_rows) == 1:
        row = candidate_rows[0]
        return {**row, "status": "행정확매칭"}

    building_tokens = tokenize_korean_terms(building_name)
    address_compact = compact_text(address_text)
    scored = []
    for record in candidate_rows:
        score = 0
        if record["extra_area"]:
            extra_compact = compact_text(record["extra_area"])
            if extra_compact and extra_compact in address_compact:
                score += 50
            for token in building_tokens:
                if token and token in extra_compact:
                    score += 20
            for token in record["extra_tokens"]:
                if token and token in address_compact:
                    score += 10
        if record["tong_ri"] and compact_text(record["tong_ri"]) in address_compact:
            score += 15
        if admin_area and record["admin_area"] == admin_area:
            score += 10
        scored.append((score, record))

    scored.sort(key=lambda item: item[0], reverse=True)
    if scored and scored[0][0] >= 40:
        return {**scored[0][1], "status": "행정확매칭"}
    if scored and scored[0][0] > 0:
        return {**scored[0][1], "status": "행후보매칭"}

    fallback = candidate_rows[0]
    return {**fallback, "status": "행후보매칭"}


def build_report(processed_students: list[dict[str, Any]], school_name: str = "", base_date: str = "") -> dict[str, Any]:
    context = get_context()
    config = load_admin_config()
    display_base_date = _format_base_date(base_date)
    aggregate: dict[tuple[str, int], dict[str, Any]] = {}
    report_records = context.report_template_records
    for record in report_records:
        aggregate[(record["sheet_name"], record["row_number"])] = {
            **record,
            **{column: 0 for column in GRADE_COLUMNS},
            "계": 0,
        }

    unmatched = []
    for student in processed_students:
        row_number = student.get("report_row_number")
        sheet_name = student.get("sheet_name")
        grade = student["grade"]

        if not row_number or (sheet_name, row_number) not in aggregate or grade not in range(1, 7):
            unmatched.append(student)
            continue

        grade_col = f"{grade}학년"
        aggregate[(sheet_name, row_number)][grade_col] += 1
        aggregate[(sheet_name, row_number)]["계"] += 1

    output = BytesIO()
    if os.path.exists(config.data_sources.report_template_source):
        workbook = load_workbook(config.data_sources.report_template_source)
        if display_base_date:
            for sheet_name in workbook.sheetnames:
                ws = workbook[sheet_name]
                if sheet_name == "재학생 현황(관내)":
                    ws["L3"] = f"기준일: {display_base_date}"
                elif sheet_name == "재학생 현황(관외)":
                    ws["M3"] = f"기준일: {display_base_date}"

        for (sheet_name, row_number), row_data in aggregate.items():
            ws = workbook[sheet_name]
            if sheet_name == "재학생 현황(관내)":
                start_col = 5
                total_col = 11
            else:
                start_col = 6
                total_col = 12
            for offset, column in enumerate(GRADE_COLUMNS):
                ws.cell(row=row_number, column=start_col + offset, value=row_data[column])
            ws.cell(row=row_number, column=total_col, value=row_data["계"])
        workbook.save(output)
    else:
        fallback_rows = list(aggregate.values())
        fallback_rows.sort(key=lambda row: (row["sheet_name"], row["row_number"]))
        fallback_df = pd.DataFrame(fallback_rows)
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            fallback_df.to_excel(writer, sheet_name="집계결과", index=False)
            if display_base_date:
                pd.DataFrame([{"기준일": display_base_date}]).to_excel(writer, sheet_name="메타", index=False)
    output.seek(0)

    table_rows = [
        row for row in aggregate.values()
        if row["category"] == "normal" and row["계"] > 0
    ]
    table_rows.sort(key=lambda row: (row["school_name"], row["admin_area"], row["tong_ri"], row["extra_area"]))

    details_df = pd.DataFrame(processed_students)
    details_output = BytesIO()
    with pd.ExcelWriter(details_output, engine="openpyxl") as writer:
        details_df.to_excel(writer, sheet_name="students", index=False)
    details_output.seek(0)

    issue_students = [
        row for row in processed_students
        if row.get("match_status") in {"학구불일치", "미분류"}
    ]
    issue_df = pd.DataFrame(issue_students)
    issue_output = BytesIO()
    with pd.ExcelWriter(issue_output, engine="openpyxl") as writer:
        issue_df.to_excel(writer, sheet_name="issues", index=False)
    issue_output.seek(0)

    summary = {
        "total_students": len(processed_students),
        "matched_students": len(processed_students) - len(unmatched),
        "unmatched_students": len(unmatched),
        "exact_matches": sum(1 for row in processed_students if row["match_status"] == "행정확매칭"),
        "candidate_matches": sum(1 for row in processed_students if row["match_status"] == "행후보매칭"),
        "school_mismatches": sum(1 for row in processed_students if row["match_status"] == "학구불일치"),
        "schools": sorted({row["school_name"] for row in table_rows if row["school_name"]}),
    }

    requested_school_workbook = b""
    if school_name:
        requested_records = [
            record for record in context.report_template_records
            if record["school_name"] == school_name
        ]
        requested_aggregate: dict[tuple[str, int], dict[str, Any]] = {}
        for record in requested_records:
            requested_aggregate[(record["sheet_name"], record["row_number"])] = {
                **record,
                **{column: 0 for column in GRADE_COLUMNS},
                "계": 0,
            }

        special_row = next(
            (
                record for record in requested_records
                if record["sheet_name"] == "재학생 현황(관내)" and record["category"] == "special" and record["admin_area"] == "기타(통학구역외)"
            ),
            None,
        )
        summary_row = next(
            (
                record for record in requested_records
                if record["sheet_name"] == "재학생 현황(관내)" and record["category"] == "summary" and record["admin_area"] == "합계"
            ),
            None,
        )

        for student in processed_students:
            grade = student["grade"]
            if grade not in range(1, 7):
                continue
            grade_col = f"{grade}학년"

            if student.get("match_status") in {"학구불일치", "미분류"}:
                if special_row:
                    bucket = requested_aggregate[(special_row["sheet_name"], special_row["row_number"])]
                    bucket[grade_col] += 1
                    bucket["계"] += 1
                continue

            if student.get("school_name") != school_name:
                continue
            row_number = student.get("report_row_number")
            sheet_name = student.get("sheet_name")
            if not row_number or (sheet_name, row_number) not in requested_aggregate:
                continue
            bucket = requested_aggregate[(sheet_name, row_number)]
            bucket[grade_col] += 1
            bucket["계"] += 1

        if summary_row:
            summary_bucket = requested_aggregate[(summary_row["sheet_name"], summary_row["row_number"])]
            normal_and_special_rows = [
                row for row in requested_aggregate.values()
                if row["sheet_name"] == summary_row["sheet_name"] and row["category"] in {"normal", "special"}
            ]
            for grade_col in GRADE_COLUMNS:
                summary_bucket[grade_col] = sum(row[grade_col] for row in normal_and_special_rows)
            summary_bucket["계"] = sum(row["계"] for row in normal_and_special_rows)

        requested_output = BytesIO()
        if os.path.exists(config.data_sources.report_template_source):
            requested_workbook = load_workbook(config.data_sources.report_template_source)
            if display_base_date:
                for sheet_name in requested_workbook.sheetnames:
                    ws = requested_workbook[sheet_name]
                    if sheet_name == "재학생 현황(관내)":
                        ws["L3"] = f"기준일: {display_base_date}"
                    elif sheet_name == "재학생 현황(관외)":
                        ws["M3"] = f"기준일: {display_base_date}"

            for (sheet_name, row_number), row_data in requested_aggregate.items():
                ws = requested_workbook[sheet_name]
                if sheet_name == "재학생 현황(관내)":
                    start_col = 5
                    total_col = 11
                else:
                    start_col = 6
                    total_col = 12
                for offset, column in enumerate(GRADE_COLUMNS):
                    ws.cell(row=row_number, column=start_col + offset, value=row_data[column])
                ws.cell(row=row_number, column=total_col, value=row_data["계"])
            requested_workbook.save(requested_output)
        else:
            requested_rows = list(requested_aggregate.values())
            requested_rows.sort(key=lambda row: (row["sheet_name"], row["row_number"]))
            requested_df = pd.DataFrame(requested_rows)
            with pd.ExcelWriter(requested_output, engine="openpyxl") as writer:
                requested_df.to_excel(writer, sheet_name="우리학교집계", index=False)
                if display_base_date:
                    pd.DataFrame([{"기준일": display_base_date}]).to_excel(writer, sheet_name="메타", index=False)
        requested_output.seek(0)
        requested_school_workbook = requested_output.getvalue()

    return {
        "summary": summary,
        "table_rows": table_rows,
        "unmatched_students": unmatched[:100],
        "report_workbook": output.getvalue(),
        "requested_school_workbook": requested_school_workbook,
        "detail_workbook": details_output.getvalue(),
        "issue_workbook": issue_output.getvalue(),
    }


RESULT_STORE: dict[str, dict[str, bytes]] = {}
RESULT_DATA_STORE: dict[str, dict[str, Any]] = {}
JOB_STORE: dict[str, dict[str, Any]] = {}
JOB_LOCK = threading.Lock()


def _now_ts() -> float:
    return time.time()


def _purge_expired_data() -> None:
    config = load_admin_config()
    now = _now_ts()
    result_ttl = max(60, config.runtime.result_ttl_minutes * 60)
    job_ttl = max(60, config.runtime.job_ttl_minutes * 60)

    with JOB_LOCK:
        expired_results = [
            result_id for result_id, payload in RESULT_DATA_STORE.items()
            if now - payload.get("created_at", now) > result_ttl
        ]
        for result_id in expired_results:
            RESULT_DATA_STORE.pop(result_id, None)
            RESULT_STORE.pop(result_id, None)

        expired_jobs = [
            job_id for job_id, payload in JOB_STORE.items()
            if now - payload.get("created_at", now) > job_ttl
        ]
        for job_id in expired_jobs:
            JOB_STORE.pop(job_id, None)


def start_cleanup_worker() -> None:
    global _CLEANUP_THREAD_STARTED
    if _CLEANUP_THREAD_STARTED:
        return
    _CLEANUP_THREAD_STARTED = True

    def worker() -> None:
        while True:
            try:
                _purge_expired_data()
                sleep_seconds = max(30, load_admin_config().runtime.cleanup_interval_seconds)
            except Exception:
                sleep_seconds = 60
            time.sleep(sleep_seconds)

    thread = threading.Thread(target=worker, name="cleanup-worker", daemon=True)
    thread.start()


def _set_job_state(job_id: str, **updates) -> None:
    with JOB_LOCK:
        if job_id in JOB_STORE:
            JOB_STORE[job_id].update(updates)


def process_request(
    files: list[tuple[str, bytes]],
    school_name: str = "",
    base_date: str = "",
    progress_callback=None,
) -> dict[str, Any]:
    context = get_context()
    if progress_callback:
        progress_callback(stage="학생 파일 읽는 중", current=0, total=0)
    students = load_student_files(files)
    processed = []
    total = len(students)
    for index, student in enumerate(students, start=1):
        processed.append(classify_student(student, context, school_name=school_name))
        if progress_callback:
            progress_callback(stage="주소 변환 및 매칭 중", current=index, total=total)
    if progress_callback:
        progress_callback(stage="서식 생성 중", current=total, total=total)
    result = build_report(processed, school_name=school_name, base_date=base_date)
    result_id = uuid4().hex
    created_at = _now_ts()
    RESULT_STORE[result_id] = {
        "report": result["report_workbook"],
        "school_report": result["requested_school_workbook"],
        "details": result["detail_workbook"],
        "issues": result["issue_workbook"],
    }
    result["created_at"] = created_at
    RESULT_DATA_STORE[result_id] = result
    result["result_id"] = result_id
    if progress_callback:
        progress_callback(stage="완료", current=total, total=total)
    return result


def get_download(result_id: str, kind: str) -> bytes | None:
    _purge_expired_data()
    bucket = RESULT_STORE.get(result_id, {})
    return bucket.get(kind)


def get_result_data(result_id: str) -> dict[str, Any] | None:
    _purge_expired_data()
    return RESULT_DATA_STORE.get(result_id)


def start_job(files: list[tuple[str, bytes]], school_name: str = "", base_date: str = "") -> str:
    job_id = uuid4().hex
    created_at = _now_ts()
    with JOB_LOCK:
        JOB_STORE[job_id] = {
            "created_at": created_at,
            "status": "queued",
            "stage": "대기 중",
            "current": 0,
            "total": 0,
            "error": "",
            "result_id": "",
        }

    def progress_callback(stage: str, current: int, total: int) -> None:
        _set_job_state(
            job_id,
            status="running" if stage != "완료" else "completed",
            stage=stage,
            current=current,
            total=total,
        )

    def worker() -> None:
        try:
            result = process_request(
                files,
                school_name=school_name,
                base_date=base_date,
                progress_callback=progress_callback,
            )
            _set_job_state(
                job_id,
                status="completed",
                stage="완료",
                current=JOB_STORE[job_id]["current"],
                total=JOB_STORE[job_id]["total"],
                result_id=result["result_id"],
            )
        except Exception as exc:
            _set_job_state(job_id, status="failed", stage="실패", error=str(exc))

    thread = threading.Thread(target=worker, name=f"job-{job_id[:8]}", daemon=True)
    thread.start()
    return job_id


def get_job(job_id: str) -> dict[str, Any] | None:
    _purge_expired_data()
    with JOB_LOCK:
        job = JOB_STORE.get(job_id)
        return dict(job) if job else None
