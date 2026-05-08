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
ADMIN_CONFIG_OVERRIDE_PATH = os.path.join(RUNTIME_DIR, "admin_config_override.json")
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


def _resolve_maybe_relative(path: str) -> str:
    """비어 있으면 빈 문자열. 절대 경로면 그대로. 아니면 저장소 루트(BASE_DIR) 기준 상대 경로."""
    text = (path or "").strip()
    if not text:
        return ""
    if os.path.isabs(text):
        return os.path.normpath(text)
    return os.path.normpath(os.path.join(BASE_DIR, text))


def _read_data_source_path(key: str, data_payload: dict, fallback: str) -> str:
    """
    admin_config.json 규칙:
    - 키가 없으면 코드 기본값(fallback, 보통 절대 경로).
    - 키가 있고 빈 문자열이면 '원본 없음'(통학구역 일람표 등 생략 시 캐시 JSON만 사용).
    - 그 외에는 상대/절대 경로 문자열.
    """
    if key not in data_payload:
        return _resolve_maybe_relative(fallback)
    raw = data_payload.get(key)
    if raw is None or str(raw).strip() == "":
        return ""
    return _resolve_maybe_relative(str(raw).strip())


def _apply_env_overrides(ds: DataSources) -> DataSources:
    """배포/로컬에서 경로만 바꿀 때: PAJU_ADMIN_SOURCE 등 환경변수가 있으면 우선."""
    admin = os.getenv("PAJU_ADMIN_SOURCE", "").strip()
    template = os.getenv("PAJU_REPORT_TEMPLATE_SOURCE", "").strip()
    zones = os.getenv("PAJU_SCHOOL_ZONE_SOURCE", "").strip()
    legacy = os.getenv("PAJU_LEGACY_DISTRICT_SOURCE", "").strip()
    return DataSources(
        admin_source=_resolve_maybe_relative(admin) if admin else ds.admin_source,
        report_template_source=_resolve_maybe_relative(template) if template else ds.report_template_source,
        school_zone_source=_resolve_maybe_relative(zones) if zones else ds.school_zone_source,
        legacy_district_source=_resolve_maybe_relative(legacy) if legacy else ds.legacy_district_source,
    )


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
    ds_def = DEFAULT_CONFIG.data_sources
    data_sources = DataSources(
        admin_source=_read_data_source_path("admin_source", data_payload, ds_def.admin_source),
        report_template_source=_read_data_source_path("report_template_source", data_payload, ds_def.report_template_source),
        school_zone_source=_read_data_source_path("school_zone_source", data_payload, ds_def.school_zone_source),
        legacy_district_source=_read_data_source_path("legacy_district_source", data_payload, ds_def.legacy_district_source),
    )
    data_sources = _apply_env_overrides(data_sources)
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
    if os.path.exists(ADMIN_CONFIG_OVERRIDE_PATH):
        try:
            with open(ADMIN_CONFIG_OVERRIDE_PATH, "r", encoding="utf-8") as file:
                override = json.load(file) or {}
            payload = {
                **payload,
                "data_sources": {
                    **(payload.get("data_sources", {}) or {}),
                    **(override.get("data_sources", {}) or {}),
                },
            }
        except Exception:
            # 오버라이드가 손상된 경우에도 기본 설정으로 안전하게 동작
            pass
    return _to_config(payload)


def save_admin_config_override(data_sources: dict[str, str]) -> None:
    """
    실행 시점에 업로드한 '붙임' 파일을 반영하기 위한 런타임 오버라이드.
    - data/runtime/ 아래에 저장되며 git 추적 대상이 아니다.
    - data_sources에는 admin_source/report_template_source 등 필요한 키만 넘기면 된다.
    """
    ensure_dirs()
    with open(ADMIN_CONFIG_OVERRIDE_PATH, "w", encoding="utf-8") as file:
        json.dump({"data_sources": data_sources}, file, ensure_ascii=False, indent=2)


def clear_admin_config_override() -> None:
    ensure_dirs()
    if os.path.exists(ADMIN_CONFIG_OVERRIDE_PATH):
        os.remove(ADMIN_CONFIG_OVERRIDE_PATH)


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
