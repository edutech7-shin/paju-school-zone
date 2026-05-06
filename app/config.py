from __future__ import annotations

from dataclasses import asdict, dataclass
import json
import os


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IS_VERCEL = bool(os.getenv("VERCEL"))
DATA_DIR = os.path.join(BASE_DIR, "data")
GENERATED_DIR = os.path.join(DATA_DIR, "generated")
ADMIN_DIR = os.path.join(DATA_DIR, "admin")
RUNTIME_DIR = os.path.join("/tmp", "addressConvert", "runtime") if IS_VERCEL else os.path.join(DATA_DIR, "runtime")
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")

REPORT_DIR = os.path.join(BASE_DIR, "통학구역별 재학생 현황 보고")
ADMIN_CONFIG_PATH = os.path.join(ADMIN_DIR, "admin_config.json")
DATASET_META_PATH = os.path.join(RUNTIME_DIR, "dataset_meta.json") if IS_VERCEL else os.path.join(GENERATED_DIR, "dataset_meta.json")


@dataclass(frozen=True)
class DataSources:
    admin_source: str
    report_template_source: str
    school_zone_source: str
    legacy_district_source: str


@dataclass(frozen=True)
class RuntimeSettings:
    app_title: str
    bind_host: str
    port: int
    result_ttl_minutes: int
    job_ttl_minutes: int
    cleanup_interval_seconds: int


@dataclass(frozen=True)
class AdminConfig:
    runtime: RuntimeSettings
    data_sources: DataSources


DEFAULT_CONFIG = AdminConfig(
    runtime=RuntimeSettings(
        app_title="파주시 초등학교 통학구역별 재학생 현황 조사",
        bind_host="0.0.0.0",
        port=8000,
        result_ttl_minutes=30,
        job_ttl_minutes=60,
        cleanup_interval_seconds=300,
    ),
    data_sources=DataSources(
        admin_source=os.path.join(REPORT_DIR, "[붙임 2] 파주시 통·리·반 명칭 및 관할구역표.xlsx"),
        report_template_source=os.path.join(REPORT_DIR, "[붙임 1] 초등학교 통학구역별 재학생 현황 서식(파주).xlsx"),
        school_zone_source=os.path.join(REPORT_DIR, "2026학년도 파주시 초등학교 통학구역 일람표.xlsx"),
        legacy_district_source=os.path.join(BASE_DIR, "관할구역.xlsx"),
    ),
)


def ensure_dirs() -> None:
    if not IS_VERCEL:
        os.makedirs(DATA_DIR, exist_ok=True)
        os.makedirs(GENERATED_DIR, exist_ok=True)
        os.makedirs(ADMIN_DIR, exist_ok=True)
    os.makedirs(RUNTIME_DIR, exist_ok=True)


def _to_config(payload: dict) -> AdminConfig:
    runtime_payload = payload.get("runtime", {})
    data_payload = payload.get("data_sources", {})
    runtime = RuntimeSettings(
        app_title=str(runtime_payload.get("app_title", DEFAULT_CONFIG.runtime.app_title)),
        bind_host=str(runtime_payload.get("bind_host", DEFAULT_CONFIG.runtime.bind_host)),
        port=int(runtime_payload.get("port", DEFAULT_CONFIG.runtime.port)),
        result_ttl_minutes=int(runtime_payload.get("result_ttl_minutes", DEFAULT_CONFIG.runtime.result_ttl_minutes)),
        job_ttl_minutes=int(runtime_payload.get("job_ttl_minutes", DEFAULT_CONFIG.runtime.job_ttl_minutes)),
        cleanup_interval_seconds=int(runtime_payload.get("cleanup_interval_seconds", DEFAULT_CONFIG.runtime.cleanup_interval_seconds)),
    )
    data_sources = DataSources(
        admin_source=str(data_payload.get("admin_source", DEFAULT_CONFIG.data_sources.admin_source)),
        report_template_source=str(data_payload.get("report_template_source", DEFAULT_CONFIG.data_sources.report_template_source)),
        school_zone_source=str(data_payload.get("school_zone_source", DEFAULT_CONFIG.data_sources.school_zone_source)),
        legacy_district_source=str(data_payload.get("legacy_district_source", DEFAULT_CONFIG.data_sources.legacy_district_source)),
    )
    return AdminConfig(runtime=runtime, data_sources=data_sources)


def save_admin_config(config: AdminConfig) -> None:
    ensure_dirs()
    with open(ADMIN_CONFIG_PATH, "w", encoding="utf-8") as file:
        json.dump(asdict(config), file, ensure_ascii=False, indent=2)


def load_admin_config() -> AdminConfig:
    ensure_dirs()
    if not os.path.exists(ADMIN_CONFIG_PATH):
        save_admin_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG
    with open(ADMIN_CONFIG_PATH, "r", encoding="utf-8") as file:
        payload = json.load(file)
    return _to_config(payload)


def get_dataset_paths() -> dict[str, str]:
    return {
        "admin": os.path.join(GENERATED_DIR, "admin_district_dataset.json"),
        "report_template": os.path.join(GENERATED_DIR, "report_template_dataset.json"),
        "school_zones": os.path.join(GENERATED_DIR, "school_zone_dataset.json"),
    }


def get_dataset_meta() -> dict:
    if not os.path.exists(DATASET_META_PATH):
        return {}
    with open(DATASET_META_PATH, "r", encoding="utf-8") as file:
        return json.load(file)


def save_dataset_meta(payload: dict) -> None:
    ensure_dirs()
    with open(DATASET_META_PATH, "w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)


def source_signature(config: AdminConfig) -> dict[str, str]:
    return {
        "admin_source": config.data_sources.admin_source,
        "report_template_source": config.data_sources.report_template_source,
        "school_zone_source": config.data_sources.school_zone_source,
        "legacy_district_source": config.data_sources.legacy_district_source,
    }


@dataclass(frozen=True)
class SourceFile:
    name: str
    path: str
