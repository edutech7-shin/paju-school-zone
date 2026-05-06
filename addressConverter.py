import pandas as pd
import requests
import time
import urllib3
import re
import json
import os
import sys
import traceback
import pickle
import threading
import queue as queue_module  # 이름 변경
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler

# 간단한 .env 로더
def load_dotenv_file(env_path=None):
    if env_path is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        env_path = os.path.join(script_dir, ".env")

    if not os.path.exists(env_path):
        return

    try:
        with open(env_path, "r", encoding="utf-8") as env_file:
            for raw_line in env_file:
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue

                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")

                if key and key not in os.environ:
                    os.environ[key] = value
    except OSError:
        return

# 파이썬 3.13에서 필요한 Empty 예외 직접 임포트
try:
    from queue import Empty  # 파이썬 3.9 이하
except ImportError:
    try:
        from _queue import Empty  # 파이썬 3.13
    except ImportError:
        # 최후의 수단으로 직접 정의
        class Empty(Exception):
            pass

# 🔧 SSL 경고 비활성화
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# .env 자동 로드
load_dotenv_file()

# ✅ 주소검색 API 키는 환경변수에서 로드
API_KEY = os.getenv("ADDRESS_CONVERT_API_KEY", "").strip()

# 🛠️ 설정
class Config:
    # 프로젝트 디렉토리를 기본 작업 위치로 사용
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = os.getenv("ADDRESS_CONVERT_BASE_DIR", SCRIPT_DIR)
    
    # 파일 경로
    INPUT_FILE = os.path.join(BASE_DIR, "전체_도로명주소_지번주소_변환용.xlsx")
    OUTPUT_FILE = os.path.join(BASE_DIR, "지번주소_변환결과.xlsx")
    TEMP_FILE = os.path.join(BASE_DIR, "지번주소_변환결과_임시.xlsx")
    FAILED_FILE = os.path.join(BASE_DIR, "지번주소_변환실패목록.xlsx")
    CACHE_FILE = os.path.join(BASE_DIR, "주소변환_캐시.pkl")
    LOG_DIR = os.path.join(BASE_DIR, "주소변환_로그")
    
    # 성능 설정
    MAX_WORKERS = 4                # 동시 실행할 작업자 수
    BATCH_SIZE = 50                # 한 번에 처리할 주소 수
    SAVE_INTERVAL = 100            # 중간 저장 간격 (행 수)
    API_TIMEOUT = 5                # API 요청 타임아웃(초)
    RETRY_COUNT = 3                # 실패 시 재시도 횟수
    THREAD_WAIT_TIMEOUT = 0.1      # 스레드 대기 타임아웃
    
    # 디버깅 설정
    DEBUG_MODE = True              # 디버그 모드
    USE_CACHE = True               # 주소 캐시 사용
    SHOW_PROGRESS_INTERVAL = 10    # 진행 상황 표시 간격

# 📝 로깅 설정
class CustomLogger:
    def __init__(self, name="주소변환", log_level=logging.INFO, log_dir=Config.LOG_DIR):
        self.logger = logging.getLogger(name)
        self.logger.setLevel(log_level)
        self.logger.handlers = []  # 기존 핸들러 제거
        
        # 로그 디렉토리 생성
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # 로그 파일 경로
        log_file = os.path.join(log_dir, f"address_converter_{datetime.now().strftime('%Y%m%d')}.log")
        error_log_file = os.path.join(log_dir, f"address_converter_errors_{datetime.now().strftime('%Y%m%d')}.log")
        
        # 포맷터 설정
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s')
        
        # 콘솔 핸들러 설정
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        console_handler.setLevel(log_level)
        self.logger.addHandler(console_handler)
        
        # 파일 핸들러 설정 (10MB 크기, 5개 파일 회전)
        file_handler = RotatingFileHandler(log_file, maxBytes=10*1024*1024, backupCount=5)
        file_handler.setFormatter(formatter)
        file_handler.setLevel(log_level)
        self.logger.addHandler(file_handler)
        
        # 에러 전용 파일 핸들러 설정
        error_handler = RotatingFileHandler(error_log_file, maxBytes=10*1024*1024, backupCount=5)
        error_handler.setFormatter(formatter)
        error_handler.setLevel(logging.ERROR)
        self.logger.addHandler(error_handler)
        
        # 로깅 시작 메시지
        self.logger.info(f"로깅 시스템 초기화 완료 (레벨: {logging.getLevelName(log_level)})")
        self.logger.info(f"로그 파일 경로: {log_file}")
        self.logger.info(f"에러 로그 파일 경로: {error_log_file}")
    
    def debug(self, message):
        self.logger.debug(message)
    
    def info(self, message):
        self.logger.info(message)
    
    def warning(self, message):
        self.logger.warning(message)
    
    def error(self, message, exc_info=False):
        self.logger.error(message, exc_info=exc_info)
    
    def critical(self, message, exc_info=True):
        self.logger.critical(message, exc_info=exc_info)
    
    def exception(self, message):
        self.logger.exception(message)

# 디버그 모드 설정
log_level = logging.DEBUG if Config.DEBUG_MODE else logging.INFO
logger = CustomLogger(log_level=log_level)

# 🔄 주소 캐시 클래스
class AddressCache:
    def __init__(self, cache_file=Config.CACHE_FILE):
        self.cache_file = cache_file
        self.cache = {}
        self.hits = 0
        self.misses = 0
        self.lock = threading.Lock()
        self.load_cache()
    
    def load_cache(self):
        """캐시 파일에서 데이터 로드"""
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'rb') as f:
                    self.cache = pickle.load(f)
                logger.info(f"주소 캐시 로드 완료: {len(self.cache)}개 항목")
            except Exception as e:
                logger.error(f"캐시 로드 실패: {str(e)}")
                self.cache = {}
        else:
            logger.info("캐시 파일이 없습니다. 새 캐시를 생성합니다.")
            self.cache = {}
    
    def save_cache(self):
        """캐시 데이터를 파일에 저장"""
        try:
            with open(self.cache_file, 'wb') as f:
                pickle.dump(self.cache, f)
            logger.debug(f"캐시 저장 완료: {len(self.cache)}개 항목")
        except Exception as e:
            logger.error(f"캐시 저장 실패: {str(e)}")
    
    def get(self, road_address):
        """도로명주소에 해당하는 지번주소 조회"""
        with self.lock:
            normalized_address = self._normalize_address(road_address)
            if normalized_address in self.cache:
                self.hits += 1
                logger.debug(f"캐시 히트: '{road_address}'")
                return self.cache[normalized_address]
            else:
                self.misses += 1
                logger.debug(f"캐시 미스: '{road_address}'")
                return None
    
    def set(self, road_address, jibun_address):
        """도로명주소와 지번주소 매핑 저장"""
        with self.lock:
            normalized_address = self._normalize_address(road_address)
            self.cache[normalized_address] = jibun_address
            
            # 주기적으로 캐시 저장 (항목 100개마다)
            if len(self.cache) % 100 == 0:
                self.save_cache()
    
    def _normalize_address(self, address):
        """주소 정규화 (공백, 특수문자 제거 등)"""
        if not address:
            return ""
        # 공백 및 특수 문자 처리, 소문자 변환
        normalized = re.sub(r'\s+', ' ', str(address).strip().lower())
        return normalized
    
    def get_stats(self):
        """캐시 통계 반환"""
        with self.lock:
            total = self.hits + self.misses
            hit_rate = (self.hits / total * 100) if total > 0 else 0
            return {
                "size": len(self.cache),
                "hits": self.hits,
                "misses": self.misses,
                "hit_rate": hit_rate
            }

# 전역 캐시 객체 초기화
address_cache = AddressCache() if Config.USE_CACHE else None

# 문자열 유사도 계산 함수
def string_similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

# API 응답 유효성 검사 함수
def validate_api_response(response_data):
    """API 응답의 유효성을 검사하고 오류 처리"""
    
    # 응답이 없는 경우
    if not response_data:
        return False, "응답 데이터가 없습니다."
    
    # common 객체 확인
    common = response_data.get('results', {}).get('common', {})
    if not common:
        return False, "응답에 common 객체가 없습니다."
    
    # 상태 코드 확인
    status_code = common.get('errorCode', '')
    error_message = common.get('errorMessage', '')
    
    # 성공 코드는 0
    if status_code != "0":
        return False, f"API 오류 (코드: {status_code}): {error_message}"
    
    # 결과 수 확인
    total_count = int(common.get('totalCount', '0'))
    if total_count == 0:
        return False, "검색 결과가 없습니다."
    
    return True, f"성공 (총 {total_count}개 결과)"

# 주소 구성요소 추출 함수 개선
def extract_address_components(road_address):
    """도로명 주소에서 다양한 구성 요소 추출 - 더 정확한 지역명 추출"""
    components = {}
    
    # 기본 주소 (쉼표 또는 괄호 이전)
    base_address = re.sub(r'\([^)]*\)', '', str(road_address)).split(',')[0].strip()
    components['base_address'] = base_address
    
    # 건물명 추출 (괄호 안)
    building_match = re.search(r'\(([^,]*?)\)', str(road_address))
    components['building_name'] = building_match.group(1).strip() if building_match else ""
    
    # 시/도 추출
    sido_match = re.match(r'^(서울특별시|부산광역시|대구광역시|인천광역시|광주광역시|대전광역시|울산광역시|세종특별자치시|경기도|강원도|충청북도|충청남도|전라북도|전라남도|경상북도|경상남도|제주특별자치도)', base_address)
    components['sido'] = sido_match.group(1) if sido_match else ""
    
    # 동/읍/면 추출 - 지역명 추출 개선
    # 괄호 안의 내용 추출 (예: (야당동))
    location_match = re.search(r'\(([^)]+)\)', str(road_address))
    if location_match:
        components['location'] = location_match.group(1).strip()
    else:
        # 괄호가 없는 경우 주소에서 동/읍/면 패턴 찾기
        location_pattern = r'((?:\w+[동읍면])|(?:\w+[마을]))'
        loc_match = re.search(location_pattern, base_address)
        components['location'] = loc_match.group(1) if loc_match else ""
    
    # 상세 주소 (동/호수 등) 추출
    detail_match = re.search(r',\s*([^,]+?)(?:\s*\([^)]*\))?$', str(road_address))
    components['detail'] = detail_match.group(1).strip() if detail_match else ""
    
    # 동/호수 정보 추가 추출
    dong_match = re.search(r'(\d+)\s*동', components['detail'])
    ho_match = re.search(r'(\d+)\s*호', components['detail'])
    
    components['dong'] = dong_match.group(1) if dong_match else ""
    components['ho'] = ho_match.group(1) if ho_match else ""
    
    # 동/호수만 있고 '동'/'호' 글자가 없는 경우 처리
    if not components['dong'] and not components['ho']:
        # 숫자만 추출해서 첫 번째는 동, 두 번째는 호수로 간주
        nums = re.findall(r'\d+', components['detail'])
        if len(nums) >= 2:
            components['dong'] = nums[0]
            components['ho'] = nums[1]
    
    # 번지 추출
    number_match = re.search(r'(\d+(?:-\d+)?)(?=번지|$)', base_address)
    components['number'] = number_match.group(1) if number_match else ""
    
    return components

# 🔁 지번주소 변환 함수 수정
def get_jibeon_address(road_address, row_index=None, expected_zip=None):
    if not API_KEY:
        raise RuntimeError(
            "환경변수 ADDRESS_CONVERT_API_KEY가 설정되지 않았습니다."
        )

    # 캐시 먼저 확인
    if Config.USE_CACHE:
        cached_result = address_cache.get(road_address)
        if cached_result:
            logger.info(f"캐시에서 결과 로드: '{road_address}' → '{cached_result}'")
            return cached_result
    
    url = "https://www.juso.go.kr/addrlink/addrLinkApi.do"
    
    # 주소 구성요소 추출
    try:
        address_parts = extract_address_components(road_address)
        if Config.DEBUG_MODE:
            logger.debug(f"주소 분석 결과: {json.dumps(address_parts, ensure_ascii=False)}")
    except Exception as e:
        logger.error(f"주소 분석 중 오류 발생: {str(e)}", exc_info=Config.DEBUG_MODE)
        address_parts = {
            'base_address': str(road_address).split(',')[0].strip(),
            'building_name': "",
            'sido': "",
            'detail': "",
            'number': ""
        }
    
    # 원본 도로명주소에서 건물명과 상세정보(동/호수) 추출
    building_name = address_parts.get('building_name', '')
    detail_info = address_parts.get('detail', '')
    
    # 상세정보 포맷팅
    detail_formatted = ""
    if detail_info:
        # 동/호 포맷팅
        dong_match = re.search(r'(\d+)동', detail_info)
        ho_match = re.search(r'(\d+)호', detail_info)
        
        # 동/호 숫자만 있고 '동'/'호' 글자가 없는 경우 처리
        if not dong_match and not ho_match:
            dong_ho_match = re.search(r'(\d+)[^\d,]*(\d+)', detail_info)
            if dong_ho_match:
                dong = dong_ho_match.group(1)
                ho = dong_ho_match.group(2)
                detail_formatted = f"{dong}동 {ho}호"
            else:
                detail_formatted = detail_info
        else:
            detail_formatted = detail_info
    
    # 키워드 생성 로직 개선 - 지역정보 우선시
    keywords = []
    
    # 1. 괄호 안의 지역명 추출 (최우선 키워드)
    location_match = re.search(r'\(([^,]+)', str(road_address))
    location_name = ""
    if location_match:
        location_name = location_match.group(1).strip()
        if location_name:
            # 지역명을 포함한 기본 주소를 최우선 키워드로 추가
            keywords.append(f"{address_parts['base_address']} {location_name}")
    
    # 2. 기본 주소 (지역명 포함)
    keywords.append(address_parts['base_address'])
    
    # 3. 건물명이 있는 경우 추가 (우선순위 하향)
    if building_name and building_name != location_name:
        # '스타벅스', '맥도날드' 등의 체인점은 키워드에서 제외
        chain_stores = ['스타벅스', '맥도날드', '버거킹', '롯데리아', '베스킨라빈스', 
                        '투썸플레이스', '이디야', '빽다방', '올리브영', 'CU', 'GS25', 
                        '세븐일레븐', '미니스톱', 'paris baguette', '파리바게뜨']
        
        is_chain_store = any(chain in building_name.lower() for chain in chain_stores)
        
        if not is_chain_store:
            # 건물명은 위치가 명확할 때만 키워드로 활용
            if location_name:
                keywords.append(f"{address_parts['base_address']} {location_name} {building_name}")
            else:
                keywords.append(f"{address_parts['base_address']} {building_name}")
    
    # 4. 2차 -> II 변환 키워드
    base_keyword = keywords[0]  # 첫번째 키워드 기준
    if '2차' in base_keyword:
        keywords.append(base_keyword.replace('2차', 'II'))
        keywords.append(base_keyword.replace('2차', 'Ⅱ'))  # 전각 로마자 추가
    
    # 중복 제거 및 빈 키워드 제거
    keywords = list(dict.fromkeys([k for k in keywords if k]))
    
    # 결과 후보 및 API 통계 초기화
    candidates = []
    api_stats = {"success": 0, "fail": 0, "total_time": 0}
    
    # 각 키워드로 검색 시도
    for i, keyword in enumerate(keywords):
        logger.debug(f"검색 시도 #{i+1}: 키워드 '{keyword}'")
        
        try:
            # API 요청 및 응답 처리
            api_start_time = time.time()
            params = {
                "confmKey": API_KEY,
                "currentPage": "1",
                "countPerPage": "10",
                "keyword": keyword,
                "resultType": "json"
            }
            
            response = requests.get(url, params=params, timeout=Config.API_TIMEOUT, verify=False)
            api_time = time.time() - api_start_time
            api_stats["total_time"] += api_time
            
            # 오류 처리 및 응답 검증
            if response.status_code != 200:
                logger.warning(f"HTTP 오류: 상태 코드 {response.status_code}")
                api_stats["fail"] += 1
                continue
            
            try:
                result = response.json()
            except json.JSONDecodeError as e:
                logger.error(f"JSON 파싱 오류: {str(e)}", exc_info=Config.DEBUG_MODE)
                api_stats["fail"] += 1
                continue
            
            is_valid, message = validate_api_response(result)
            if not is_valid:
                logger.warning(f"API 응답: {message}")
                api_stats["fail"] += 1
                continue
            
            api_stats["success"] += 1
            juso_list = result.get('results', {}).get('juso', [])
            
            if not juso_list:
                logger.warning(f"검색 키워드 '{keyword}'에 대한 결과가 없습니다.")
                continue
            
            # 점수 계산 로직 개선
            for juso in juso_list:
                score = 0
                temp_scores = {}
                
                # 1. 도로명주소 유사도 (0-40점) - 가중치 상향 조정
                road_similarity = string_similarity(address_parts['base_address'].lower(), juso['roadAddr'].lower())
                temp_scores['road_score'] = road_similarity * 40
                
                # 2. 지역명 일치 여부 (0-30점) - 가중치 상향 조정
                temp_scores['location_score'] = 0
                if location_name:
                    # 지번주소에 지역명이 포함되었는지 확인 (예: 야당동)
                    if location_name in juso['jibunAddr']:
                        temp_scores['location_score'] = 30
                    elif juso.get('emdNm') and location_name in juso['emdNm']:
                        temp_scores['location_score'] = 25
                
                # 3. 건물명 일치 여부 (0-20점) - 가중치 하향 조정
                temp_scores['building_score'] = 0
                if building_name and juso.get('bdNm'):
                    # 체인점 상호명은 매칭에서 제외
                    chain_stores = ['스타벅스', '맥도날드', '버거킹', '롯데리아', '베스킨라빈스', 
                                    '투썸플레이스', '이디야', '빽다방', '올리브영', 'CU', 'GS25', 
                                    '세븐일레븐', '미니스톱', 'paris baguette', '파리바게뜨']
                    
                    is_chain_store = any(chain in juso['bdNm'].lower() for chain in chain_stores)
                    
                    if not is_chain_store:
                        building_similarity = string_similarity(
                            building_name.lower(), 
                            juso['bdNm'].lower()
                        )
                        temp_scores['building_score'] = building_similarity * 20
                        
                        # 건물명 완전 일치 시 보너스 점수 (10점)
                        if building_name.lower() == juso['bdNm'].lower():
                            temp_scores['building_bonus'] = 10
                        else:
                            temp_scores['building_bonus'] = 0
                else:
                    temp_scores['building_bonus'] = 0
                
                # 4. 행정구역 일치도 (0-10점)
                temp_scores['admin_score'] = 0
                if address_parts['sido'] and juso.get('siNm'):
                    admin_similarity = string_similarity(
                        address_parts['sido'], 
                        juso['siNm'] + ' ' + juso.get('sggNm', '')
                    )
                    temp_scores['admin_score'] = admin_similarity * 10

                # 5. 우편번호 일치 여부 (0-25점)
                temp_scores['zip_score'] = 0
                if expected_zip and juso.get('zipNo'):
                    if str(expected_zip).strip() == str(juso.get('zipNo')).strip():
                        temp_scores['zip_score'] = 25
                
                # 총점 계산
                score = sum(temp_scores.values())
                
                # 중복 제거 로직
                duplicate = False
                for c in candidates:
                    if c['jibunAddr'] == juso['jibunAddr']:
                        if score > c['score']:
                            c['score'] = score
                            c['keyword'] = keyword
                            c['details'] = temp_scores
                        duplicate = True
                        break
                
                if not duplicate:
                    candidates.append({
                        'jibunAddr': juso['jibunAddr'],
                        'roadAddr': juso['roadAddr'],
                        'bdNm': juso.get('bdNm', ''),
                        'zipNo': juso.get('zipNo', ''),
                        'score': score,
                        'keyword': keyword,
                        'details': temp_scores
                    })
            
            logger.debug(f"키워드 '{keyword}'로 {len(juso_list)}개 결과 발견")
            
            # 고품질 결과 검색 조건 개선 - 지역명 일치를 우선시
            high_quality_match = False
            for c in candidates:
                # 지역명이 일치하고 점수가 높은 경우
                location_score = c['details'].get('location_score', 0)
                if location_score > 20 and c['score'] > 60:
                    high_quality_match = True
                    break
            
            if len(candidates) >= 5 and high_quality_match:
                logger.debug("지역명이 일치하는 고품질 결과를 찾았습니다. 추가 검색 중단.")
                break
                
        except requests.exceptions.RequestException as e:
            logger.error(f"HTTP 요청 오류: {str(e)}", exc_info=Config.DEBUG_MODE)
            api_stats["fail"] += 1
            time.sleep(1)
        except Exception as e:
            logger.error(f"예상치 못한 오류: {str(e)}", exc_info=Config.DEBUG_MODE)
            api_stats["fail"] += 1
            time.sleep(0.5)
    
    # 모든 시도 후 결과 선택 로직 개선
    if candidates:
        # 점수 기준 내림차순 정렬
        candidates.sort(key=lambda x: x['score'], reverse=True)
        
        # 상위 후보들 로그 출력
        if Config.DEBUG_MODE:
            logger.debug(f"🎯 상위 후보 결과 ({len(candidates)}개 중):")
            for i, c in enumerate(candidates[:5], 1):
                detail_scores = ", ".join([f"{k.split('_')[0]}:{v:.1f}" for k, v in c['details'].items()])
                logger.debug(f"{i}. 점수: {c['score']:.1f} ({detail_scores}), 키워드: '{c['keyword']}'")
                logger.debug(f"   지번: {c['jibunAddr']}, 건물명: {c['bdNm']}")
        
        # 체인점 상호명 필터링 - 선택 로직 개선
        chain_stores = ['스타벅스', '맥도날드', '버거킹', '롯데리아', '베스킨라빈스', 
                        '투썸플레이스', '이디야', '빽다방', '올리브영', 'CU', 'GS25', 
                        '세븐일레븐', '미니스톱', 'paris baguette', '파리바게뜨']
        
        filtered_candidates = []
        for candidate in candidates:
            is_chain_store = False
            if candidate.get('bdNm'):
                is_chain_store = any(chain in candidate['bdNm'].lower() for chain in chain_stores)
            
            if not is_chain_store:
                filtered_candidates.append(candidate)
        
        # 필터링 후 후보가 있으면 해당 후보에서 선택, 없으면 원본 후보에서 선택
        selection_pool = filtered_candidates if filtered_candidates else candidates
        
        # 지역명 일치 우선 선택 로직
        location_matches = []
        if location_name:
            for candidate in selection_pool:
                if location_name in candidate['jibunAddr']:
                    location_matches.append(candidate)
        
        # 최종 선택
        if location_matches:
            # 지역명 일치하는 후보 중 최고 점수 선택
            location_matches.sort(key=lambda x: x['score'], reverse=True)
            selected_candidate = location_matches[0]
            logger.debug(f"지역명 '{location_name}'이 일치하는 후보를 선택했습니다.")
        else:
            # 지역명 일치하는 후보가 없으면 최고 점수 선택
            selected_candidate = selection_pool[0]
        
        # 선택된 지번주소 처리
        selected_jibun = selected_candidate['jibunAddr']
        
        # 원본 상세정보(동/호수)가 있으면 지번주소에 추가
        if detail_formatted:
            if not any(part in selected_jibun.lower() for part in detail_formatted.lower().split()):
                if not selected_jibun.endswith((',', ' ')):
                    selected_jibun += ', '
                selected_jibun += detail_formatted
                logger.debug(f"상세정보 '{detail_formatted}'를 지번주소에 추가했습니다.")
        
        # 괄호 제거 - 주소 끝에 있는 닫는 괄호 제거
        selected_jibun = selected_jibun.rstrip(')')
        
        logger.info(f"✓ 최종 선택 지번주소: {selected_jibun}")
        
        # 캐시에 결과 저장
        if Config.USE_CACHE:
            address_cache.set(road_address, selected_jibun)
        
        return selected_jibun
    
    # 결과가 전혀 없는 경우
    logger.error(f"모든 검색 시도 실패: '{road_address}'")
    return None

# 작업자 스레드 함수
def worker_thread(task_queue, results, progress):
    """작업자 스레드 - 주소 변환 작업 처리"""
    thread_name = threading.current_thread().name
    logger.debug(f"{thread_name} 시작")
    
    while True:
        try:
            # 큐에서 작업 가져오기 (타임아웃 적용)
            task = task_queue.get(timeout=Config.THREAD_WAIT_TIMEOUT)
            
            # 종료 신호 확인
            if task is None:
                logger.debug(f"{thread_name} 종료 신호 수신")
                task_queue.task_done()
                break
            
            idx, road_address = task
            
            try:
                # 지번주소 변환
                jibun_address = get_jibeon_address(road_address, idx)
                
                # 결과 저장
                with results.lock:
                    results.data[idx] = jibun_address
                    results.processed += 1
                
                # 진행 상황 업데이트
                with progress.lock:
                    if jibun_address:
                        progress.success += 1
                    else:
                        progress.fail += 1
                
            except Exception as e:
                logger.exception(f"작업 처리 중 오류: {str(e)}")
                with progress.lock:
                    progress.fail += 1
            
            # 작업 완료 표시
            task_queue.task_done()
            
        except Exception:
            # 큐가 비었을 때 - 잠시 대기 후 다시 시도
            # 파이썬 3.13에서의 Empty 예외 호환성 문제 해결을 위해 일반 Exception 사용
            continue
    
    logger.debug(f"{thread_name} 종료")

# 결과 저장 객체
class Results:
    def __init__(self):
        self.data = {}  # 인덱스 -> 지번주소 매핑
        self.processed = 0  # 처리된 작업 수
        self.lock = threading.Lock()

# 진행 상황 객체
class Progress:
    def __init__(self, total):
        self.total = total  # 총 작업 수
        self.success = 0  # 성공 작업 수
        self.fail = 0  # 실패 작업 수
        self.last_save = 0  # 마지막 저장 시점
        self.start_time = time.time()  # 시작 시간
        self.lock = threading.Lock()
    
    def get_progress(self):
        """진행률 0.0-1.0 반환"""
        completed = self.success + self.fail
        return completed / self.total if self.total > 0 else 0
    
    def get_stats(self):
        """진행 통계 정보 반환"""
        completed = self.success + self.fail
        elapsed = time.time() - self.start_time
        
        # 예상 남은 시간 계산
        progress = self.get_progress()
        if progress > 0:
            estimated_total = elapsed / progress
            remaining = estimated_total - elapsed
        else:
            remaining = 0
        
        # 성공률 계산
        success_rate = (self.success / completed * 100) if completed > 0 else 0
        
        return {
            "total": self.total,
            "completed": completed,
            "success": self.success,
            "fail": self.fail,
            "progress": progress * 100,  # 백분율
            "elapsed": elapsed,
            "remaining": remaining,
            "success_rate": success_rate
        }

# 변환 진행률 표시 함수
def show_progress(progress):
   """진행 상황 표시"""
   stats = progress.get_stats()
   
   # 진행 막대 생성 (폭 30자)
   bar_width = 30
   bar = '█' * int(bar_width * stats["progress"]/100) + '░' * (bar_width - int(bar_width * stats["progress"]/100))
   
   # 시간 정보
   elapsed_str = time.strftime("%H:%M:%S", time.gmtime(stats["elapsed"]))
   remaining_str = time.strftime("%H:%M:%S", time.gmtime(stats["remaining"]))
   
   # 진행률 출력
   logger.info(f"진행률: |{bar}| {stats['completed']}/{stats['total']} ({stats['progress']:.1f}%)")
   logger.info(f"성공: {stats['success']}건 ({stats['success_rate']:.1f}%), 실패: {stats['fail']}건")
   logger.info(f"시간: 경과 {elapsed_str}, 남음 {remaining_str}")
   
   # 캐시 통계 (캐시 사용 시)
   if Config.USE_CACHE:
       cache_stats = address_cache.get_stats()
       logger.info(f"캐시: {cache_stats['size']}항목, 히트율 {cache_stats['hit_rate']:.1f}%")

# 중간 저장 함수
def save_progress(df, results, filename=Config.TEMP_FILE):
   """현재까지의 결과를 엑셀 파일로 저장"""
   try:
       # 결과 데이터프레임에 반영
       for idx, jibun_addr in results.data.items():
           # 데이터 타입 문제 방지를 위해 명시적 타입 변환
           if jibun_addr is not None:
               df.at[idx, '지번주소'] = str(jibun_addr)
       
       # 파일 저장
       df.to_excel(filename, index=False)
       logger.info(f"중간 결과 저장 완료: {filename} ({results.processed}건)")
       return True
   except Exception as e:
       logger.error(f"중간 결과 저장 실패: {str(e)}", exc_info=True)
       return False

# 배치 처리 함수
def process_in_batches(df, start_idx=0):
   """데이터프레임의 주소를 배치 단위로 처리"""
   # 처리할 행 필터링 (지번주소가 없는 행)
   to_process_df = df[df['지번주소'].isna() | (df['지번주소'] == 'nan')].iloc[start_idx:]
   total_rows = len(to_process_df)
   
   if total_rows == 0:
       logger.info("처리할 주소가 없습니다.")
       return True
   
   logger.info(f"처리할 주소: {total_rows}개 (전체 {len(df)}개 중)")
   
   # 재개 지점 로그
   if start_idx > 0:
       logger.info(f"이전 작업에서 재개: 인덱스 {start_idx}부터 시작")
   
   # 작업 큐 초기화
   task_queue = queue_module.Queue()
   
   # 결과 및 진행 상황 객체 초기화
   results = Results()
   progress = Progress(total_rows)
   
   # 작업자 스레드 생성
   workers = []
   for i in range(Config.MAX_WORKERS):
       worker = threading.Thread(
           target=worker_thread,
           args=(task_queue, results, progress),
           name=f"Worker-{i+1}"
       )
       worker.daemon = True
       worker.start()
       workers.append(worker)
       logger.debug(f"작업자 스레드 {i+1} 시작")
   
   try:
       # 행 처리 루프
       batch_count = 0
       for batch_start in range(0, total_rows, Config.BATCH_SIZE):
           batch_end = min(batch_start + Config.BATCH_SIZE, total_rows)
           batch_df = to_process_df.iloc[batch_start:batch_end]
           batch_count += 1
           
           logger.info(f"배치 {batch_count} 처리 시작 (행 {batch_start+1}-{batch_end})")
           
           # 작업 큐에 배치 추가
           for idx, row in batch_df.iterrows():
               task_queue.put((idx, row['도로명주소']))
           
           # 배치 완료 대기
           task_queue.join()
           logger.info(f"배치 {batch_count} 처리 완료")
           
           # 진행 상황 표시
           if batch_end % Config.SHOW_PROGRESS_INTERVAL == 0 or batch_end == total_rows:
               show_progress(progress)
           
           # 중간 저장
           if batch_end - progress.last_save >= Config.SAVE_INTERVAL or batch_end == total_rows:
               if save_progress(df, results):
                   progress.last_save = batch_end
               
               # 캐시 저장 (캐시 사용 시)
               if Config.USE_CACHE:
                   address_cache.save_cache()
       
       # 모든 작업 완료
       logger.info("모든 배치 처리 완료")
       return True
       
   except KeyboardInterrupt:
       logger.warning("사용자에 의한 작업 중단")
       return False
   except Exception as e:
       logger.critical(f"배치 처리 중 예상치 못한 오류: {str(e)}", exc_info=True)
       return False
   finally:
       # 작업자 스레드 종료
       logger.debug("작업자 스레드 종료 중...")
       for _ in range(len(workers)):
           task_queue.put(None)  # 종료 신호 전송
       
       # 스레드 조인 (최대 5초 대기)
       for worker in workers:
           worker.join(5)
       
       # 작업 통계
       stats = progress.get_stats()
       logger.info(f"작업 통계: 완료 {stats['completed']}/{stats['total']} ({stats['progress']:.1f}%)")
       logger.info(f"성공: {stats['success']}건, 실패: {stats['fail']}건, 성공률: {stats['success_rate']:.1f}%")
       logger.info(f"소요 시간: {time.strftime('%H:%M:%S', time.gmtime(stats['elapsed']))}")
       
       # 캐시 저장
       if Config.USE_CACHE:
           address_cache.save_cache()
           cache_stats = address_cache.get_stats()
           logger.info(f"캐시 통계: {cache_stats['size']}항목, 히트율 {cache_stats['hit_rate']:.1f}%")
       
       # 중간 저장
       save_progress(df, results)

# 작업 재개 지점 찾기 함수
def find_resume_point(df):
   """이전 작업의 재개 지점 찾기"""
   # 이미 변환된 행 수 계산
   processed = df['지번주소'].notna() & (df['지번주소'] != 'nan')
   processed_count = processed.sum()
   
   # 전체 행 수
   total = len(df)
   
   # 미처리 행이 있는지 확인
   if processed_count < total:
       logger.info(f"이전 작업 발견: {processed_count}/{total} 행 처리됨")
       return processed_count
   else:
       logger.info("모든 행이 이미 처리되었습니다.")
       return 0

# 메인 함수
def main():
   """메인 프로그램 실행"""
   logger.info("=" * 80)
   logger.info("도로명주소 → 지번주소 변환 프로그램 시작 (멀티스레드 최적화 버전)")
   logger.info(f"버전: 2.2.0 (스타벅스 버그 수정)")
   logger.info(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
   logger.info(f"설정: 작업자={Config.MAX_WORKERS}, 배치크기={Config.BATCH_SIZE}, 캐시사용={Config.USE_CACHE}")
   logger.info("=" * 80)
   
   try:
       # 📥 엑셀 파일 불러오기
       logger.info(f"엑셀 파일 로드 중: {Config.INPUT_FILE}")
       
       # 파일 존재 확인
       if not os.path.exists(Config.INPUT_FILE):
           logger.critical(f"파일을 찾을 수 없습니다: {Config.INPUT_FILE}")
           return False
       
       try:
           # 데이터 타입 문제 해결을 위해 명시적으로 문자열 타입 지정
           df = pd.read_excel(Config.INPUT_FILE, dtype={'지번주소': str})
           logger.info(f"파일 로드 성공: {len(df)}개 행")
       except Exception as e:
           logger.critical(f"엑셀 파일 로드 오류: {str(e)}", exc_info=True)
           return False
       
       # 필수 컬럼 확인
       if '도로명주소' not in df.columns:
           logger.critical("필수 컬럼 '도로명주소'가 파일에 없습니다.")
           return False
       
       # 지번주소 컬럼 추가 (없는 경우)
       if '지번주소' not in df.columns:
           logger.info("'지번주소' 컬럼을 추가합니다.")
           df['지번주소'] = pd.NA  # None 대신 pd.NA 사용
       
       # 작업 재개 지점 찾기
       resume_point = find_resume_point(df)
       
       # 배치 처리 실행
       success = process_in_batches(df, resume_point)
       
       # 최종 결과 저장
       logger.info("최종 결과 저장 중...")
       df.to_excel(Config.OUTPUT_FILE, index=False)
       logger.info(f"✅ 변환 완료! 저장 경로: {Config.OUTPUT_FILE}")
       
       # 실패 목록 저장
       failed_df = df[df['지번주소'].isna() | (df['지번주소'] == 'nan')]
       if len(failed_df) > 0:
           failed_df.to_excel(Config.FAILED_FILE, index=False)
           logger.info(f"❗ 변환 실패 목록 저장: {Config.FAILED_FILE} ({len(failed_df)}건)")
       
       # 성공 여부 반환
       return success
       
   except KeyboardInterrupt:
       # 사용자에 의한 중단 처리
       logger.warning("\n작업이 사용자에 의해 중단되었습니다.")
       return False
       
   except Exception as e:
       # 예상치 못한 오류 처리
       logger.critical("프로그램 실행 중 치명적 오류 발생:", exc_info=True)
       return False

# 프로그램 실행
if __name__ == "__main__":
   try:
       success = main()
       exit_code = 0 if success else 1
       
       # 종료 메시지
       logger.info("=" * 80)
       logger.info(f"프로그램 종료 (상태: {'성공' if success else '실패'})")
       logger.info("=" * 80)
       
       sys.exit(exit_code)
       
   except Exception as e:
       logger.critical(f"프로그램 종료 처리 중 오류 발생: {str(e)}", exc_info=True)
       sys.exit(1)
   finally:
       # 캐시 저장
       if Config.USE_CACHE and address_cache:
           address_cache.save_cache()
