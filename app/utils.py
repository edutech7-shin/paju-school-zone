from __future__ import annotations

import math
import re
from typing import Iterable


def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip() == ""


def clean_text(value) -> str:
    if is_blank(value):
        return ""
    return str(value).replace("\xa0", " ").strip()


def compact_text(value) -> str:
    text = clean_text(value)
    return re.sub(r"\s+", "", text)


def normalize_range_symbols(value: str) -> str:
    return (
        clean_text(value)
        .replace("～", "~")
        .replace("∼", "~")
        .replace("〜", "~")
        .replace("ㆍ", "·")
    )


def tokenize_korean_terms(value: str) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    parts = re.findall(r"[가-힣A-Za-z0-9]+", text)
    return [part for part in parts if len(part) >= 2]


def any_contains(text: str, candidates: Iterable[str]) -> bool:
    normalized = compact_text(text)
    return any(token and token in normalized for token in candidates)


def parse_jibun_address(address: str) -> dict:
    text = normalize_range_symbols(address)
    match = re.search(r"([가-힣]+(?:동|리))\s*(산)?\s*(\d+)(?:-(\d+))?", text)
    if not match:
        return {
            "legal_area": "",
            "is_mountain": False,
            "main_number": None,
            "sub_number": 0,
        }
    return {
        "legal_area": match.group(1),
        "is_mountain": bool(match.group(2)),
        "main_number": int(match.group(3)),
        "sub_number": int(match.group(4) or 0),
    }


def extract_parenthetical_buildings(text: str) -> list[str]:
    buildings: list[str] = []
    for content in re.findall(r"\(([^)]+)\)", clean_text(text)):
        for item in content.split(","):
            token = clean_text(item)
            if token:
                buildings.append(token)
    return buildings


def strip_parenthetical(text: str) -> str:
    return re.sub(r"\([^)]*\)", "", clean_text(text)).strip()


def parse_address_unit_info(text: str) -> dict:
    normalized = normalize_range_symbols(clean_text(text))
    block_numbers = [match.group(1) for match in re.finditer(r"(\d+)동", normalized)]
    unit_numbers = [int(match.group(1)) for match in re.finditer(r"(\d+)호", normalized)]
    return {
        "block_numbers": block_numbers,
        "unit_numbers": unit_numbers,
    }


def parse_block_unit_ranges(text: str) -> tuple[list[str], list[dict]]:
    normalized = normalize_range_symbols(clean_text(text))
    block_numbers: list[str] = []
    unit_ranges: list[dict] = []
    current_block = ""
    last_chunk_was_unit = False

    for chunk in [part.strip() for part in normalized.split(",") if part.strip()]:
        block_match = re.search(r"(\d+)동", chunk)
        if block_match:
            current_block = block_match.group(1)
            if current_block not in block_numbers:
                block_numbers.append(current_block)

        matched_unit = False
        for unit_match in re.finditer(r"(\d+)(?:~(\d+))?호", chunk):
            start = int(unit_match.group(1))
            end = int(unit_match.group(2) or unit_match.group(1))
            unit_ranges.append(
                {
                    "block_number": current_block,
                    "start_unit": start,
                    "end_unit": end,
                    "raw": unit_match.group(0),
                }
            )
            matched_unit = True
        if not matched_unit and current_block and last_chunk_was_unit:
            bare_unit_match = re.fullmatch(r"(\d+)(?:~(\d+))?", chunk.replace(" ", ""))
            if bare_unit_match:
                start = int(bare_unit_match.group(1))
                end = int(bare_unit_match.group(2) or bare_unit_match.group(1))
                unit_ranges.append(
                    {
                        "block_number": current_block,
                        "start_unit": start,
                        "end_unit": end,
                        "raw": chunk,
                    }
                )
                matched_unit = True
        last_chunk_was_unit = matched_unit
    return block_numbers, unit_ranges


def unit_in_ranges(unit_number: int | None, ranges: list[dict], block_number: str = "") -> bool:
    if unit_number is None:
        return False

    for item in ranges:
        range_block = clean_text(item.get("block_number", ""))
        if block_number and range_block and range_block != block_number:
            continue
        if item["start_unit"] <= unit_number <= item["end_unit"]:
            return True
    return False


def parse_parcel_ranges(text: str, legal_area: str = "") -> list[dict]:
    cleaned = normalize_range_symbols(strip_parenthetical(text))
    if legal_area:
        cleaned = cleaned.replace(legal_area, " ")
    raw_tokens = [token.strip() for token in cleaned.split(",") if token.strip()]
    ranges: list[dict] = []
    tokens: list[str] = []
    current_main: int | None = None
    current_is_mountain = False
    last_token_had_sub = False
    current_block = ""
    last_chunk_was_unit = False

    for raw_token in raw_tokens:
        block_match = re.search(r"(\d+)동", raw_token)
        if block_match:
            current_block = block_match.group(1)
        if "호" in raw_token:
            last_chunk_was_unit = True
            continue
        if current_block and last_chunk_was_unit and re.fullmatch(r"\d+(?:~\d+)?", raw_token.replace(" ", "")):
            continue
        token = raw_token.replace(" ", "")
        token = re.sub(r"[^0-9~\-산]", "", token)
        if not token:
            continue
        if token.startswith("-") and current_main is not None:
            token = f"{'산' if current_is_mountain else ''}{current_main}{token}"
        elif (
            re.fullmatch(r"\d+", token)
            and current_main is not None
            and last_token_had_sub
            and len(token) < len(str(current_main))
        ):
            token = f"{'산' if current_is_mountain else ''}{current_main}-{token}"
        current_match = re.match(r"(산)?(\d+)", token)
        if current_match:
            current_is_mountain = bool(current_match.group(1))
            current_main = int(current_match.group(2))
        last_token_had_sub = "-" in token.split("~", 1)[0]
        last_chunk_was_unit = False
        tokens.append(token)

    for token in tokens:
        token = token.replace(" ", "")

        start_token, end_token = (token.split("~", 1) + [None])[:2]

        def parse_side(side: str):
            match = re.fullmatch(r"(산)?(\d+)(?:-(\d+))?", side)
            if not match:
                return None
            return {
                "is_mountain": bool(match.group(1)),
                "main": int(match.group(2)),
                "sub": int(match.group(3) or 0),
                "has_sub": match.group(3) is not None,
            }

        start = parse_side(start_token)
        if not start:
            continue

        if end_token is None:
            end = {
                "is_mountain": start["is_mountain"],
                "main": start["main"],
                "sub": start["sub"],
            }
        else:
            end = parse_side(end_token)
            if not end:
                continue
            if "-" not in end_token and start["has_sub"]:
                end = {
                    "is_mountain": start["is_mountain"],
                    "main": start["main"],
                    "sub": int(end_token.replace("산", "")),
                }
            elif "-" not in end_token and not start["has_sub"]:
                end = {
                    "is_mountain": start["is_mountain"],
                    "main": int(end_token.replace("산", "")),
                    "sub": 0,
                }

        ranges.append(
            {
                "is_mountain": start["is_mountain"],
                "start_main": start["main"],
                "start_sub": start["sub"],
                "end_main": end["main"],
                "end_sub": end["sub"],
                "end_is_mountain": end["is_mountain"],
                "raw": token,
            }
        )

    return ranges


def parcel_in_ranges(parcel: dict, ranges: list[dict]) -> bool:
    if parcel.get("main_number") is None:
        return False

    for item in ranges:
        if parcel.get("is_mountain", False) != item["is_mountain"]:
            continue
        main = parcel["main_number"]
        sub = parcel.get("sub_number", 0)
        start_main = item["start_main"]
        end_main = item["end_main"]
        start_sub = item["start_sub"]
        end_sub = item["end_sub"]

        if start_main == end_main:
            if main != start_main:
                continue
            if start_sub <= sub <= end_sub:
                return True
            continue

        if main < start_main or main > end_main:
            continue
        if main == start_main and sub < start_sub:
            continue
        if main == end_main and sub > end_sub:
            continue
        return True

    return False
