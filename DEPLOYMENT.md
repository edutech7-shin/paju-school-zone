# 배포 안내 (로컬 + Render)

파주시 초등학교 통학구역별 재학생 현황 조사 웹앱 배포 방법입니다.

## 1) 로컬/내부망 실행

- 실행: `python3 webapp.py`
- 기본 접속: `http://0.0.0.0:8000`

## 2) Render 배포 구조

이 저장소는 다음 배포형 구조를 사용합니다.

- `api/index.py`: Render 웹 프로세스 진입점(Flask app)
- `Procfile`: Render Start Command 정의
- `requirements.txt`: Python 의존성 설치

## 3) 환경변수

`.env.example` 기준으로 아래 값을 Render Environment Variables에 등록합니다.

- `ADDRESS_CONVERT_API_KEY`: 도로명주소 API 키
- `SUPABASE_URL`: Supabase 프로젝트 URL
- `SUPABASE_ANON_KEY`: Supabase anon key

> 현재 앱 핵심 처리(주소 변환/분류/엑셀 생성)는 서버 메모리 기반이며, Supabase 키는 배포 환경 표준화 용도로 먼저 반영했습니다.

## 4) Render 배포 명령

```bash
# Render Dashboard에서 GitHub repo 연결 후 자동 배포
# Start Command: gunicorn api.index:app --bind 0.0.0.0:$PORT
```

## GitHub 업로드 전 정리

- `.gitignore`로 로그/캐시/대용량 엑셀 산출물 제외
- 배포에 필요한 코드/템플릿/정적 파일(`app/`, `api/`, `templates/`, `static/`, `data/generated/`)만 유지

## 5) 설정 관리

운영 중 설정 변경은 코드/환경변수 기준으로 반영합니다.

- `data/admin/admin_config.json`으로 기본 설정 관리
- `REPORT_TEMPLATE_URL` 환경변수로 공용 보고서 서식 파일 URL 관리

## 6) 개인정보 보호

- 업로드 파일은 처리 중 메모리에서만 사용
- 결과/작업 기록은 메모리 저장 후 TTL 자동 삭제
- 기본 TTL
  - `result_ttl_minutes`: 30분
  - `job_ttl_minutes`: 60분
  - `cleanup_interval_seconds`: 300초

## 7) 운영 시 주의

- 로컬 내부망 운영 시 외부 공개 금지
- Render 인스턴스 재시작 시 메모리 데이터는 유지되지 않음
