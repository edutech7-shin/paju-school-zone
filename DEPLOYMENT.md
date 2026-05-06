# 배포 안내 (로컬 + Vercel)

파주시 초등학교 통학구역별 재학생 현황 조사 웹앱 배포 방법입니다.

## 1) 로컬/내부망 실행

- 실행: `python3 webapp.py`
- 기본 접속: `http://0.0.0.0:8000`
- 관리자 화면: `http://127.0.0.1:8000/admin`

## 2) Vercel 배포 구조

이 저장소는 다음 배포형 구조를 사용합니다.

- `api/index.py`: Vercel Python Serverless 진입점
- `vercel.json`: 모든 요청을 `api/index.py`로 라우팅
- `.vercelignore`: 배포에 불필요한 로그/산출물/대용량 파일 제외
- `requirements.txt`: Vercel 빌드 시 Python 의존성 설치

## 3) 환경변수

`.env.example` 기준으로 아래 값을 Vercel Project Environment Variables에 등록합니다.

- `ADDRESS_CONVERT_API_KEY`: 도로명주소 API 키
- `SUPABASE_URL`: Supabase 프로젝트 URL
- `SUPABASE_ANON_KEY`: Supabase anon key

> 현재 앱 핵심 처리(주소 변환/분류/엑셀 생성)는 서버 메모리 기반이며, Supabase 키는 배포 환경 표준화 용도로 먼저 반영했습니다.

## 4) Vercel 배포 명령

```bash
vercel login
vercel link
vercel --prod
```

## GitHub 업로드 전 정리

- `.gitignore`로 로그/캐시/대용량 엑셀 산출물 제외
- `.vercelignore`로 Vercel 배포 아카이브에서 불필요 파일 제외
- 배포에 필요한 코드/템플릿/정적 파일(`app/`, `api/`, `templates/`, `static/`, `data/generated/`)만 유지

## 5) 관리자 설정 항목

`/admin` 또는 `data/admin/admin_config.json`에서 아래를 수정할 수 있습니다.

- 프로그램 제목
- 서버 바인드 주소/포트
- 결과 자동 삭제 시간
- 작업 기록 자동 삭제 시간
- 통·리·반 원본 파일 경로
- 보고서 서식 파일 경로
- 학교 통학구역 파일 경로
- 보조 관할구역 파일 경로
  - `legacy_district_source`는 선택 항목이며, 비어 있거나 파일이 없어도 앱은 동작합니다.

## 6) 개인정보 보호

- 업로드 파일은 처리 중 메모리에서만 사용
- 결과/작업 기록은 메모리 저장 후 TTL 자동 삭제
- 기본 TTL
  - `result_ttl_minutes`: 30분
  - `job_ttl_minutes`: 60분
  - `cleanup_interval_seconds`: 300초

## 7) 운영 시 주의

- 로컬 내부망 운영 시 외부 공개 금지
- 관리자 화면은 별도 인증이 없으므로 접근 통제 필수
- Vercel Serverless 특성상 인스턴스 재시작 시 메모리 데이터는 유지되지 않음
