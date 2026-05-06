import pandas as pd
import numpy as np
import os
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

# 학년 숫자 추출 함수
def extract_grade_number(grade_str):
    if pd.isna(grade_str):
        return None
    match = re.search(r'(\d+)', str(grade_str))
    return int(match.group(1)) if match else None

# 주소에서 동이름, 본번, 부번 추출 함수
def extract_address_components(address):
    if pd.isna(address):
        return None, None, None
    match = re.search(r'([가-힣]+동)\s*(\d+)(?:-(\d+))?', str(address))
    if match:
        return match.group(1), int(match.group(2)), int(match.group(3) or 0)
    return None, None, None

# 건물명 추출 함수 개선
def extract_building_name(address):
    if pd.isna(address):
        return None
    address_str = str(address)
    
    # 도로명이 포함된 주소에서 건물명 추출
    # 패턴: 도로명+번호 다음에 오는 괄호 안의 내용 (야당동,건물명)
    road_building_pattern = r'(송학길|하우고개길|번뛰기길)\s+\d+.*?\((야당동|동패동),\s*([^)]+)\)'
    road_match = re.search(road_building_pattern, address_str)
    if road_match:
        return road_match.group(3)  # 괄호 안의 건물명 반환
    
    # 일반 건물명 패턴 추출
    building_pattern = r'([가-힣0-9]+(?:아파트|빌라|단지|타운|파크|시티|하우스|레지던스|팰리스|힐스|마을|빌|하임|1차|2차|3차))'
    matches = re.findall(building_pattern, address_str)
    
    # 특수 케이스 처리: 롯데캐슬파크타운 2차/Ⅱ
    lotte_pattern_regular = re.search(r'롯데캐슬.*파크타운.*2차', address_str)
    lotte_pattern_roman = re.search(r'롯데캐슬.*파크타운.*Ⅱ', address_str)
    
    if lotte_pattern_regular or lotte_pattern_roman:
        matches.append("롯데캐슬파크타운 2차")
    elif '롯데캐슬' in address_str and '파크타운' in address_str:
        matches.append("롯데캐슬파크타운")
    
    # 특수 건물명 패턴 추가
    special_patterns = [
        r'한빛마을\s*3단지', r'한빛마을\s*4단지', r'한빛마을\s*9단지',
        r'자연에\s*빌라', r'해드림아트', r'자연에가\s*3',
        r'로얄클래스', r'롯데캐슬', r'자유로아이파크', r'리치타운',
        r'휴하우스', r'메디하임', r'라포레빌'
    ]
    
    for pattern in special_patterns:
        match = re.search(pattern, address_str)
        if match:
            matches.append(match.group(0))
    
    # 가장 긴 건물명 선택
    if matches:
        return sorted(set(matches), key=len, reverse=True)[0]
    return None

# 주소 표준화 함수 개선
def standardize_address(address):
    if pd.isna(address):
        return None
    
    address_str = str(address)
    
    # 도로명 주소 처리 (송학길, 하우고개길, 번뛰기길)
    road_pattern = r'(송학길|하우고개길|번뛰기길)\s+(\d+)'
    road_match = re.search(road_pattern, address_str)
    if road_match:
        road_name = road_match.group(1)
        road_number = road_match.group(2)
        return f"{road_name} {road_number}"
    
    # 지번 주소 처리
    jibun_match = re.search(r'([가-힣]+동)\s*(\d+)(?:-(\d+))?', address_str)
    if jibun_match:
        dong = jibun_match.group(1)
        main = jibun_match.group(2)
        sub = jibun_match.group(3)
        return f"{dong} {main}" + (f"-{sub}" if sub else "")
    
    return None

# 주소 분석 및 추출 함수 개선
def analyze_address(address):
    """주소에서 도로명, 동이름, 본번, 부번 등을 추출하는 함수"""
    if pd.isna(address):
        return {
            'road_name': None,
            'road_number': None,
            'dong_name': None,
            'main_number': None,
            'sub_number': None,
            'is_road_address': False
        }
    
    address_str = str(address)
    result = {
        'road_name': None,
        'road_number': None,
        'dong_name': None,
        'main_number': None,
        'sub_number': None,
        'is_road_address': False
    }
    
    # 도로명 주소 추출
    road_pattern = r'(송학길|하우고개길|번뛰기길)\s+(\d+)'
    road_match = re.search(road_pattern, address_str)
    if road_match:
        result['road_name'] = road_match.group(1)
        result['road_number'] = int(road_match.group(2))
        result['is_road_address'] = True
        
        # 도로명 주소에서 동이름 추출 (괄호 안에 있는 경우가 많음)
        dong_pattern = r'\((야당동|동패동)'
        dong_match = re.search(dong_pattern, address_str)
        if dong_match:
            result['dong_name'] = dong_match.group(1)
    
    # 지번 주소 추출
    jibun_match = re.search(r'([가-힣]+동)\s*(\d+)(?:-(\d+))?', address_str)
    if jibun_match:
        result['dong_name'] = jibun_match.group(1)
        result['main_number'] = int(jibun_match.group(2))
        result['sub_number'] = int(jibun_match.group(3) or 0)
    
    return result

# 통 이름 정리 함수
def clean_tong_name(tong):
    if pd.isna(tong):
        return tong
    return re.sub(r'^제', '', str(tong))

# 관할구역 전처리 개선 함수
def preprocess_district_data(district_df):
    """관할구역 데이터를 전처리하여 번지 범위를 더 정확하게 추출하는 함수"""
    
    # 결과를 저장할 리스트
    processed_districts = []
    
    for _, row in district_df.iterrows():
        dong = row['동']
        tong = row['통']
        area_desc = str(row['관할구역']) if not pd.isna(row['관할구역']) else ""
        
        # 동 이름 추출
        dong_names = re.findall(r'([가-힣]+동)', area_desc)
        
        if not dong_names:
            # 동 이름이 없으면 그대로 보존
            processed_districts.append({
                '동': dong,
                '통': tong,
                '관할구역': area_desc,
                '동이름': None,
                '번지범위': None
            })
            continue
        
        # 각 동 이름별로 처리
        for dong_name in dong_names:
            # 해당 동 주변 텍스트 추출
            dong_pattern = f'{dong_name}[^가-힣]*?([^가-힣동]+?)(?:[가-힣]+동|$)'
            dong_match = re.search(dong_pattern, area_desc)
            
            if dong_match:
                dong_area = dong_match.group(1).strip()
            else:
                dong_area = ""
            
            # 번지 범위 추출
            ranges = []
            
            # 1. 범위 패턴 (123~456, 123-45~678-90 등)
            range_pattern = r'(\d+)(?:-(\d+))?\s*[~～]\s*(\d+)(?:-(\d+))?'
            for match in re.finditer(range_pattern, dong_area):
                start_main = int(match.group(1))
                start_sub = int(match.group(2) or 0)
                end_main = int(match.group(3))
                end_sub = int(match.group(4) or 0)
                
                ranges.append({
                    '유형': '범위',
                    '시작본번': start_main,
                    '시작부번': start_sub,
                    '끝본번': end_main,
                    '끝부번': end_sub
                })
            
            # 2. 단일 번지 패턴 (123, 123-45 등)
            single_pattern = r'(?<![~～\d-])(\d+)(?:-(\d+))?(?![~～\d-])'
            for match in re.finditer(single_pattern, dong_area):
                main = int(match.group(1))
                sub = int(match.group(2) or 0)
                
                # 이미 범위에 포함된 번지는 제외
                if not any(r['유형'] == '범위' and 
                          r['시작본번'] <= main <= r['끝본번'] for r in ranges):
                    ranges.append({
                        '유형': '단일',
                        '본번': main,
                        '부번': sub
                    })
            
            # 추출된 범위를 정리하여 추가
            processed_districts.append({
                '동': dong,
                '통': tong,
                '관할구역': area_desc,
                '동이름': dong_name,
                '번지범위': ranges
            })
    
    # 데이터프레임으로 변환하여 반환
    return pd.DataFrame(processed_districts)

# 번지수 범위 추출 함수 개선
def parse_district_ranges(district_df):
    district_ranges = []
    for _, row in district_df.iterrows():
        dong = row['동']
        tong = row['통']
        area_desc = row['관할구역']
        
        if pd.isna(area_desc):
            continue
        
        # 동 이름 추출 개선 - 여러 동이 있을 경우도 처리
        all_dongs = re.findall(r'([가-힣]+동)', str(area_desc))
        if not all_dongs:
            continue
        
        for area_dong in all_dongs:
            # 해당 동 주변의 텍스트 추출하여 번지 범위 찾기
            dong_text = re.search(f'{area_dong}[^가-힣]*([^가-힣동]+)', str(area_desc))
            if dong_text:
                dong_area = dong_text.group(1)
            else:
                dong_area = area_desc
            
            # 범위 패턴 추출 (예: 123-45~678-90) - 더 유연한 패턴 인식
            range_pattern = r'(\d+)(?:-(\d+))?\s*[~～]\s*(\d+)(?:-(\d+))?'
            for match in re.finditer(range_pattern, str(dong_area)):
                district_ranges.append({
                    '동이름': area_dong,
                    '관할동': dong,
                    '관할통': tong,
                    '시작본번': int(match.group(1)),
                    '시작부번': int(match.group(2) or 0),
                    '끝본번': int(match.group(3)),
                    '끝부번': int(match.group(4) or 0)
                })
            
            # 단일 번지 패턴 추출 (예: 123-45)
            # 연속된 숫자(-숫자)들을 개별 번지로 인식
            single_pattern = r'(?<![~～\d-])(\d+)(?:-(\d+))?(?![~～\d-])'
            for match in re.finditer(single_pattern, str(dong_area)):
                main = int(match.group(1))
                sub = int(match.group(2) or 0)
                
                # 이미 범위에 포함된 번지는 제외
                if not any(r['동이름'] == area_dong and 
                          r['시작본번'] <= main <= r['끝본번'] for r in district_ranges):
                    district_ranges.append({
                        '동이름': area_dong,
                        '관할동': dong,
                        '관할통': tong,
                        '시작본번': main,
                        '시작부번': sub,
                        '끝본번': main,
                        '끝부번': sub
                    })
    
    return pd.DataFrame(district_ranges)

# 관할구역 데이터에서 지번주소 매핑 테이블 생성 함수
def create_jibun_mapping_table(processed_district_df):
    """관할구역 데이터를 기반으로 상세한 지번주소 매핑 테이블 생성"""
    
    mapping_table = {}
    
    for _, district in processed_district_df.iterrows():
        dong = district['동']
        tong = str(district['통']).replace('제', '') if not pd.isna(district['통']) else '미확인'
        dong_name = district['동이름']
        ranges = district['번지범위']
        
        if pd.isna(dong_name) or not ranges:
            continue
        
        # 범위에 속하는 모든 가능한 지번주소 생성
        for range_info in ranges:
            if range_info['유형'] == '범위':
                # 본번 범위 순회
                for main in range(range_info['시작본번'], range_info['끝본번'] + 1):
                    # 시작과 끝 번호의 부번 처리
                    if main == range_info['시작본번'] and main == range_info['끝본번']:
                        # 시작과 끝 본번이 같은 경우, 부번 범위 처리
                        for sub in range(range_info['시작부번'], range_info['끝부번'] + 1):
                            if sub == 0:
                                key = f"{dong_name} {main}"
                            else:
                                key = f"{dong_name} {main}-{sub}"
                            mapping_table[key] = (dong, tong)
                    elif main == range_info['시작본번']:
                        # 시작 본번인 경우, 시작부번부터 확인
                        for sub in range(range_info['시작부번'], 100):  # 부번 상한 설정
                            if sub == 0:
                                key = f"{dong_name} {main}"
                            else:
                                key = f"{dong_name} {main}-{sub}"
                            mapping_table[key] = (dong, tong)
                    elif main == range_info['끝본번']:
                        # 끝 본번인 경우, 끝부번까지 확인
                        for sub in range(0, range_info['끝부번'] + 1):
                            if sub == 0:
                                key = f"{dong_name} {main}"
                            else:
                                key = f"{dong_name} {main}-{sub}"
                            mapping_table[key] = (dong, tong)
                    else:
                        # 중간 본번인 경우, 모든 부번 포함
                        key = f"{dong_name} {main}"
                        mapping_table[key] = (dong, tong)
                        # 일반적인 부번 범위 (0~99)도 추가
                        for sub in range(1, 100):
                            key = f"{dong_name} {main}-{sub}"
                            mapping_table[key] = (dong, tong)
            
            elif range_info['유형'] == '단일':
                # 단일 번지 처리
                main = range_info['본번']
                sub = range_info['부번']
                
                if sub == 0:
                    key = f"{dong_name} {main}"
                else:
                    key = f"{dong_name} {main}-{sub}"
                
                mapping_table[key] = (dong, tong)
    
    return mapping_table

# 야당동 번지 매핑 테이블 구성 함수
def create_yadang_mapping_table():
    """야당동 주요 번지와 통을 매핑하는 테이블 생성 함수"""
    
    yadang_mapping = {}
    
    # 운정3동에 속하는 야당동 번지
    # 한빛마을 3단지 주변 (야당동 970~990 번대)
    for main in range(970, 990):
        yadang_mapping[f"야당동 {main}"] = ('운정3동', '18-19통')
        for sub in range(1, 20):
            yadang_mapping[f"야당동 {main}-{sub}"] = ('운정3동', '18-19통')
    
    # 한빛마을 9단지 주변 (야당동 1010~1020 번대)
    for main in range(1010, 1020):
        if main != 1015:  # 1015는 롯데캐슬파크타운 2차
            yadang_mapping[f"야당동 {main}"] = ('운정3동', '30통')
            for sub in range(1, 20):
                yadang_mapping[f"야당동 {main}-{sub}"] = ('운정3동', '30통')
    
    # 롯데캐슬파크타운 2차 (야당동 1015)
    yadang_mapping["야당동 1015"] = ('운정3동', '31통')
    for sub in range(1, 20):
        yadang_mapping[f"야당동 1015-{sub}"] = ('운정3동', '31통')
    
    # 운정4동에 속하는 야당동 번지 (예시, 실제 데이터로 조정 필요)
    # 180~190 번대 (휴하우스 주변)
    for main in range(180, 190):
        yadang_mapping[f"야당동 {main}"] = ('운정4동', '12통')
        for sub in range(1, 50):
            yadang_mapping[f"야당동 {main}-{sub}"] = ('운정4동', '12통')
    
    # 다른 번대 구역들도 추가...
    mapping_ranges = [
        {'range': (200, 250), 'dong': '운정4동', 'tong': '10통'},
        {'range': (250, 300), 'dong': '운정4동', 'tong': '11통'},
        {'range': (300, 350), 'dong': '운정4동', 'tong': '8통'},
        {'range': (350, 400), 'dong': '운정4동', 'tong': '9통'},
        {'range': (400, 450), 'dong': '운정4동', 'tong': '13통'},
        {'range': (450, 500), 'dong': '운정4동', 'tong': '15통'},
        {'range': (500, 550), 'dong': '운정4동', 'tong': '16통'},
        {'range': (550, 600), 'dong': '운정4동', 'tong': '17통'},
        {'range': (600, 650), 'dong': '운정4동', 'tong': '18통'},
    ]
    
    for range_info in mapping_ranges:
        start, end = range_info['range']
        dong = range_info['dong']
        tong = range_info['tong']
        
        for main in range(start, end):
            yadang_mapping[f"야당동 {main}"] = (dong, tong)
            for sub in range(1, 50):  # 일반적인 부번 범위
                yadang_mapping[f"야당동 {main}-{sub}"] = (dong, tong)
    
    return yadang_mapping

# 동 별 통 관할구역 정리 함수
def summarize_districts_by_dong(district_df):
    """각 동별로 통과 관할구역을 정리하여 보여주는 함수"""
    
    summary = {}
    
    for _, row in district_df.iterrows():
        dong = row['동']
        tong = row['통']
        area_desc = row['관할구역']
        
        if dong not in summary:
            summary[dong] = []
        
        # 관할구역 정보가 있는 경우만 추가
        if not pd.isna(area_desc):
            summary[dong].append({'통': tong, '관할구역': area_desc})
    
    return summary

# 번지수 기반 매핑 함수 개선 - 중요: 이 함수가 map_address_to_school_area 함수보다 먼저 정의되어야 함
def map_address_by_jibun(dong, main_num, sub_num, district_ranges_df):
    if pd.isna(dong) or pd.isna(main_num):
        return None, None
    
    # 해당 동에 속하는 모든 범위 데이터 필터링
    matched_districts = district_ranges_df[district_ranges_df['동이름'] == dong]
    
    # 매칭된 구역이 없으면 None 반환
    if matched_districts.empty:
        return None, None
    
    # 정확한 매칭 (본번과 부번이 모두 범위 내)
    for _, district_range in matched_districts.iterrows():
        start_main = district_range['시작본번']
        start_sub = district_range['시작부번']
        end_main = district_range['끝본번']
        end_sub = district_range['끝부번']
        
        # 본번이 범위 내에 있는 경우
        if start_main <= main_num <= end_main:
            # 1. 본번이 시작과 끝이 같은 경우 (예: 123 ~ 123)
            if main_num == start_main and main_num == end_main:
                # 부번도 범위 내에 있어야 함
                if start_sub <= sub_num <= end_sub:
                    return district_range['관할동'], district_range['관할통']
            # 2. 본번이 시작 번호와 같은 경우
            elif main_num == start_main:
                # 부번이 시작 부번 이상이어야 함
                if sub_num >= start_sub:
                    return district_range['관할동'], district_range['관할통']
            # 3. 본번이 끝 번호와 같은 경우
            elif main_num == end_main:
                # 부번이 끝 부번 이하여야 함
                if sub_num <= end_sub:
                    return district_range['관할동'], district_range['관할통']
            # 4. 본번이 범위 내에 있고, 시작/끝 번호와 다른 경우
            else:
                return district_range['관할동'], district_range['관할통']
    
    # 가장 가까운 번지수 찾기 (정확한 매칭이 없을 경우)
    # 같은 본번에 가장 가까운 부번, 또는 가장 가까운 본번 찾기
    closest_match = None
    min_distance = float('inf')
    
    for _, district_range in matched_districts.iterrows():
        start_main = district_range['시작본번']
        end_main = district_range['끝본번']
        
        # 주변 본번 찾기
        if abs(main_num - start_main) < min_distance:
            min_distance = abs(main_num - start_main)
            closest_match = (district_range['관할동'], district_range['관할통'])
        
        if abs(main_num - end_main) < min_distance:
            min_distance = abs(main_num - end_main)
            closest_match = (district_range['관할동'], district_range['관할통'])
    
    # 임계값 설정 (본번 차이가 너무 크면 다른 지역일 가능성이 높음)
    if min_distance <= 20:  # 본번 차이가 20 이내인 경우만 매핑
        return closest_match
    
    return None, None

# 야당동 번지 범위에 따른 통 매핑을 위한 개선된 함수
def map_yadang_address_by_range(main_num, sub_num):
    """
    야당동의 번지 범위를 기반으로 통을 매핑하는 함수
    """
    # 야당동 주요 번지 범위 및 통 매핑 (실제 관할구역 데이터 기반으로 보완 필요)
    yadang_ranges = [
        # 운정3동 구역
        {'range': (970, 990), 'dong': '운정3동', 'tong': '18-19통'},  # 한빛마을 3단지 주변
        {'range': (1010, 1020), 'dong': '운정3동', 'tong': '30통'},  # 한빛마을 9단지
        {'range': (1015, 1015), 'dong': '운정3동', 'tong': '31통'},  # 롯데캐슬파크타운 2차
        
        # 운정4동 구역 (범위는 예시로 실제 데이터로 조정 필요)
        {'range': (180, 190), 'dong': '운정4동', 'tong': '12통'},
        {'range': (200, 250), 'dong': '운정4동', 'tong': '10통'},
        {'range': (250, 300), 'dong': '운정4동', 'tong': '11통'},
        {'range': (300, 350), 'dong': '운정4동', 'tong': '8통'},
        {'range': (350, 400), 'dong': '운정4동', 'tong': '9통'},
        {'range': (400, 450), 'dong': '운정4동', 'tong': '13통'},
        {'range': (450, 500), 'dong': '운정4동', 'tong': '15통'},
        {'range': (500, 550), 'dong': '운정4동', 'tong': '16통'},
        {'range': (550, 600), 'dong': '운정4동', 'tong': '17통'},
        {'range': (600, 650), 'dong': '운정4동', 'tong': '18통'},
    ]
    
    if pd.isna(main_num):
        return None, None
    
    # 번지 범위 내에 있는지 확인
    for range_info in yadang_ranges:
        start, end = range_info['range']
        if start <= main_num <= end:
            return range_info['dong'], range_info['tong']
    
    return None, None

  # 지번주소 테스트 함수
def test_address_mapping(address, district_ranges_df, jibun_mapping_table, manual_mapping):
    """특정 주소의 매핑 결과를 테스트하고 결과를 보여주는 함수"""
    
    print(f"테스트 주소: {address}")
    
    # 주소 성분 추출
    dong, main_num, sub_num = extract_address_components(address)
    building = extract_building_name(address)
    std_address = standardize_address(address)
    
    print(f"추출된 정보:")
    print(f"  - 동이름: {dong}")
    print(f"  - 본번: {main_num}")
    print(f"  - 부번: {sub_num}")
    print(f"  - 건물명: {building}")
    print(f"  - 표준화된 주소: {std_address}")
    
    # 매핑 결과 확인
    result = None
    source = None
    
    # 1. 매뉴얼 매핑 확인
    if building and building in manual_mapping:
        result = manual_mapping[building]
        source = "매뉴얼 매핑 (건물명)"
    elif std_address and std_address in manual_mapping:
        result = manual_mapping[std_address]
        source = "매뉴얼 매핑 (표준주소)"
    elif dong and main_num:
        key = f"{dong} {main_num}"
        if key in manual_mapping:
            result = manual_mapping[key]
            source = "매뉴얼 매핑 (본번)"
        elif key in jibun_mapping_table:
            result = jibun_mapping_table[key]
            source = "지번주소 매핑 테이블"
        else:
            # 2. 지번주소 부번 매핑 확인
            if sub_num:
                sub_key = f"{dong} {main_num}-{sub_num}"
                if sub_key in jibun_mapping_table:
                    result = jibun_mapping_table[sub_key]
                    source = "지번주소 매핑 테이블 (부번 포함)"
            
            # 3. 범위 매핑 시도
            if not result and dong and main_num is not None:
                dong_result, tong_result = map_address_by_jibun(dong, main_num, sub_num, district_ranges_df)
                if dong_result and tong_result:
                    result = (dong_result, tong_result)
                    source = "번지수 범위 기반 매핑"
                # 야당동 특별 처리
                elif dong == '야당동':
                    range_dong, range_tong = map_yadang_address_by_range(main_num, sub_num)
                    if range_dong and range_tong:
                        result = (range_dong, range_tong)
                        source = "야당동 번지 범위 매핑"
    
    # 결과 출력
    if result:
        print(f"\n매핑 결과: {result[0]}, {result[1]}")
        print(f"매핑 소스: {source}")
    else:
        print("\n매핑 결과: 없음 (미확인)")
    
    return result, source

# 번지별 통계 분석 함수
def analyze_district_mapping_coverage(students_df, district_ranges_df):
    """학생 데이터와 관할구역 데이터의 매핑 커버리지를 분석하는 함수"""
    
    # 동별 본번 집계
    dong_main_counts = {}
    
    for _, row in students_df.iterrows():
        dong = row['dong_name']
        main_num = row['main_number']
        
        if pd.isna(dong) or pd.isna(main_num):
            continue
        
        if dong not in dong_main_counts:
            dong_main_counts[dong] = {}
        
        if main_num not in dong_main_counts[dong]:
            dong_main_counts[dong][main_num] = 0
        
        dong_main_counts[dong][main_num] += 1
    
    # 관할구역 데이터 분석
    district_coverage = {}
    
    for dong in dong_main_counts:
        # 해당 동에 속하는 관할구역 추출
        dong_districts = district_ranges_df[district_ranges_df['동이름'] == dong]
        
        if dong_districts.empty:
            district_coverage[dong] = {
                '총 번지 수': len(dong_main_counts[dong]),
                '매핑된 번지 수': 0,
                '매핑 비율': 0
            }
            continue
        
        # 매핑 가능한 번지 확인
        mappable_mains = []
        all_mains = list(dong_main_counts[dong].keys())
        
        for main in all_mains:
            for _, district in dong_districts.iterrows():
                start_main = district['시작본번']
                end_main = district['끝본번']
                
                if start_main <= main <= end_main:
                    mappable_mains.append(main)
                    break
        
        # 매핑 통계 계산
        district_coverage[dong] = {
            '총 번지 수': len(all_mains),
            '매핑된 번지 수': len(mappable_mains),
            '매핑 비율': len(mappable_mains) / len(all_mains) if all_mains else 0
        }
    
    return district_coverage, dong_main_counts

# 수동 매핑 추가
manual_mapping = {
    # 한빛마을 관련
    '한빛마을3단지': ('운정3동', '18-19통'),
    '한빛마을 3단지': ('운정3동', '18-19통'),
    '한빛마을4단지': ('운정3동', '31통'),
    '한빛마을 4단지': ('운정3동', '31통'),
    '한빛마을9단지': ('운정3동', '30통'),
    '한빛마을 9단지': ('운정3동', '30통'),
    
    # 롯데캐슬 관련
    '롯데캐슬파크타운 2차': ('운정3동', '31통'),
    '롯데캐슬파크타운2차': ('운정3동', '31통'),
    '롯데캐슬 파크타운 2차': ('운정3동', '31통'),
    '롯데캐슬 파크타운 Ⅱ': ('운정3동', '31통'),
    '롯데캐슬 파크타운Ⅱ': ('운정3동', '31통'),
    '롯데캐슬파크타운Ⅱ': ('운정3동', '31통'),
    '롯데캐슬파크타운 Ⅱ': ('운정3동', '31통'),
    '롯데캐슬파크타운': ('운정3동', '29통'),
    
    # 번지수 기반 특별 매핑
    '야당동 1015': ('운정3동', '31통'),  # 롯데캐슬파크타운 2차
    '야당동 1016': ('운정3동', '30통'),  # 한빛마을 9단지
    '야당동 975': ('운정3동', '18-19통'),
    '야당동 979': ('운정3동', '30통'),
    
    # 새로 추가: 휴하우스 및 야당동의 다른 건물들
    '휴하우스': ('운정4동', '12통'),
    '야당동 183-36': ('운정4동', '12통'),  # 휴하우스 지번
    
    # 도로명 기반 매핑
    '송학길': ('운정4동', '미확인'),
    '하우고개길': ('운정4동', '미확인'),
    '번뛰기길': ('운정4동', '미확인'),
    
    # 특정 도로명 주소 번지 (예시 - 실제 통 번호는 확인 필요)
    '번뛰기길 8': ('운정4동', '12통'),
    '번뛰기길 10': ('운정4동', '12통'),
    '하우고개길 30': ('운정4동', '15통'),
    '송학길 36': ('운정4동', '13통'),
}

# 주소 매핑 함수 개선
def map_address_to_school_area(row, manual_mapping, district_building_map, district_address_map, district_ranges_df):
    address = row['지번주소']
    std_address = row['std_address']
    building = row['building_name']
    dong = row['dong_name']
    main_num = row['main_number']
    sub_num = row['sub_number']
    
    if pd.isna(address):
        return ('미확인', '미확인')
    
    # 1. 매뉴얼 매핑 (건물명 기반) - 가장 우선 적용
    if pd.notna(building) and building in manual_mapping:
        return manual_mapping[building]
    
    # 2. 매뉴얼 매핑 (표준화된 주소 기반)
    if pd.notna(std_address) and std_address in manual_mapping:
        return manual_mapping[std_address]
    
    # 3. district_building_map 확인 (사전 정의된 건물명 매핑)
    if pd.notna(building) and building in district_building_map:
        return district_building_map[building]
    
    # 4. district_address_map 확인 (사전 정의된 주소 매핑)
    if pd.notna(std_address) and std_address in district_address_map:
        return district_address_map[std_address]
    
    # 야당동 특별 처리
    if dong == '야당동':
        # 지정된 특수 케이스 먼저 처리
        if main_num == 1015:  # 롯데캐슬파크타운 2차
            return ('운정3동', '31통')
        elif main_num == 1016:  # 한빛마을 9단지
            return ('운정3동', '30통')
        elif main_num == 975 or main_num == 974:  # 한빛마을 3단지
            return ('운정3동', '18-19통')
        elif main_num == 979:  # 한빛마을 9단지 관련
            return ('운정3동', '30통')
        
        # 건물명 기반 특수 케이스
        if pd.notna(building):
            if '한빛마을3' in building or '한빛마을 3' in building:
                return ('운정3동', '18-19통')
            elif '한빛마을4' in building or '한빛마을 4' in building:
                return ('운정3동', '31통')
            elif '한빛마을9' in building or '한빛마을 9' in building:
                return ('운정3동', '30통')
            elif '롯데캐슬' in building and '파크타운' in building:
                if '2차' in building or 'Ⅱ' in building:
                    return ('운정3동', '31통')
                else:
                    return ('운정3동', '29통')
        
        # 번지 기반 야당동 매핑
        if pd.notna(main_num):
            # 관할구역 번지 기반 매핑 시도
            dong_result, tong_result = map_address_by_jibun(dong, main_num, sub_num, district_ranges_df)
            if dong_result and tong_result:
                tong_result = re.sub(r'^제', '', str(tong_result))
                return (dong_result, tong_result)
            
            # 야당동 내의 번지 범위 기반 매핑 시도
            range_dong, range_tong = map_yadang_address_by_range(main_num, sub_num)
            if range_dong and range_tong:
                return (range_dong, range_tong)
    
    # 동패동 및 다른 동 처리
    if pd.notna(dong) and pd.notna(main_num):
        # 번지 기반 매핑 시도
        dong_result, tong_result = map_address_by_jibun(dong, main_num, sub_num, district_ranges_df)
        if dong_result and tong_result:
            tong_result = re.sub(r'^제', '', str(tong_result))
            return (dong_result, tong_result)
    
    # 주소 텍스트 기반 매핑
    address_str = str(address).lower()
    
    # 특정 도로명을 포함한 주소 처리
    if '송학길' in address_str or '하우고개길' in address_str or '번뛰기길' in address_str:
        # 도로명+번호 패턴 추출
        road_pattern = r'(송학길|하우고개길|번뛰기길)\s+(\d+)'
        road_match = re.search(road_pattern, address_str)
        
        if road_match:
            road_name = road_match.group(1)
            road_number = road_match.group(2)
            road_key = f"{road_name} {road_number}"
            
            # 해당 도로명+번호가 manual_mapping에 있는지 확인
            if road_key in manual_mapping:
                return manual_mapping[road_key]
            
            # 번뛰기길 특정 번호 처리
            if road_name == '번뛰기길':
                if road_number == '8':
                    return ('운정4동', '12통')
            
            # 괄호 안에 있는 동-번지 추출 시도
            dong_jibun_pattern = r'\((야당동|동패동)[^)]*?(\d+)(?:-(\d+))?\)'
            dong_jibun_match = re.search(dong_jibun_pattern, address_str)
            
            if dong_jibun_match:
                road_dong = dong_jibun_match.group(1)
                road_main = int(dong_jibun_match.group(2))
                road_sub = int(dong_jibun_match.group(3) or 0)
                
                # 추출된 지번으로 매핑 시도
                dong_result, tong_result = map_address_by_jibun(road_dong, road_main, road_sub, district_ranges_df)
                if dong_result and tong_result:
                    tong_result = re.sub(r'^제', '', str(tong_result))
                    return (dong_result, tong_result)
        
        # 특정 번호 매핑이 없는 경우
        # 도로명 주소를 사용하되, 기본 통은 확인 못한 것으로 처리
        if '야당동' in address_str:
            return ('운정4동', '미확인')
        if '동패동' in address_str:
            return ('운정4동', '미확인')
    
    # 주소 텍스트에서 패턴 기반 매핑
    if '롯데캐슬' in address_str and '파크타운' in address_str and ('2차' in address_str or 'ⅱ' in address_str.lower()):
        return ('운정3동', '31통')
    if '한빛마을9' in address_str or '한빛마을 9' in address_str or '야당동 979' in address_str:
        return ('운정3동', '30통')
    if '야당동 1016' in address_str:
        return ('운정3동', '30통')
    if '야당동 975' in address_str or ('야당동' in address_str and '한빛마을3' in address_str):
        return ('운정3동', '18-19통')
    
    # 매뉴얼 매핑에서 주소 텍스트로 검색
    for building_name, area in manual_mapping.items():
        if building_name.lower() in address_str:
            return area
    
    # 기본 동 매핑 (최후의 수단)
    if dong == '야당동':
        # 명확한 번지 정보가 없으면서 특별 케이스에 해당하지 않을 때만
        # 번지수가 있으면서 매핑되지 않으면 운정4동으로 기본 매핑
        return ('운정4동', '미확인')
    elif dong == '동패동':
        return ('운정4동', '미확인')
    elif '운정3동' in address_str:
        return ('운정3동', '미확인')
    elif '운정4동' in address_str:
        return ('운정4동', '미확인')
    elif '경기도 파주시' in address_str:
        return ('파주시', '미확인')
    
    return ('기타', '통학구역외')

# 결과 포맷팅 함수
def create_formatted_result(result_df, output_dir):
    # 새 워크북 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "통학구역별 학생수"
    
    # 제목 행 추가
    ws.merge_cells('A1:B1')
    ws['A1'] = '통학구역'
    ws.merge_cells('C1:I1')
    ws['C1'] = '학생수'
    
    # 헤더 행 추가
    headers = ['동', '통', '기타', '1학년', '2학년', '3학년', '4학년', '5학년', '6학년', '기타학년', '계']
    for col, header in enumerate(headers, 1):
        ws.cell(row=2, column=col, value=header)
    
    # 데이터 채우기
    for row_idx, row in result_df.iterrows():
        if (row['동'] == '기타' and row['통'] == '통학구역외'):
            ws.cell(row=row_idx+3, column=1, value='')
            ws.cell(row=row_idx+3, column=2, value='')
            ws.cell(row=row_idx+3, column=3, value='통학구역외')
        elif row['통'] == '미확인':
            if row['동'] == '파주시':
                ws.cell(row=row_idx+3, column=1, value='')
                ws.cell(row=row_idx+3, column=2, value='')
                ws.cell(row=row_idx+3, column=3, value='파주시 미확인')
            else:
                ws.cell(row=row_idx+3, column=1, value=row['동'])
                ws.cell(row=row_idx+3, column=2, value='')
                ws.cell(row=row_idx+3, column=3, value='미확인')
        else:
            ws.cell(row=row_idx+3, column=1, value=row['동'])
            ws.cell(row=row_idx+3, column=2, value=row['통'])
            ws.cell(row=row_idx+3, column=3, value='')
        
        # 학년별 학생 수
        for grade in range(1, 7):
            col_name = f'{grade}학년'
            ws.cell(row=row_idx+3, column=grade+3, value=row.get(col_name, 0))
        
        # 기타 학년 학생 수
        ws.cell(row=row_idx+3, column=10, value=row.get('기타학년', 0))
        
        # 합계
        ws.cell(row=row_idx+3, column=11, value=row.get('계', 0))
    
    # 마지막 행 인덱스
    last_row = len(result_df) + 2
    
    # 합계 행 추가
    ws.cell(row=last_row+1, column=1, value='합계')
    ws.merge_cells(f'A{last_row+1}:C{last_row+1}')
    
    # 열별 합계 계산
    for col in range(4, 12):
        sum_formula = f'=SUM({chr(64+col)}3:{chr(64+col)}{last_row})'
        ws.cell(row=last_row+1, column=col, value=sum_formula)
    
    # 스타일 적용
    light_blue_fill = PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    
    for row in [1, 2]:
        for col in range(1, 12):
            cell = ws.cell(row=row, column=col)
            cell.fill = light_blue_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(bold=True)
    
    for row in range(3, last_row+1):
        for col in range(4, 12):
            cell = ws.cell(row=row, column=col)
            cell.fill = yellow_fill
            cell.alignment = Alignment(horizontal='center')
    
    for col in range(1, 12):
        cell = ws.cell(row=last_row+1, column=col)
        cell.alignment = Alignment(horizontal='center')
        cell.font = Font(bold=True)
    
    # 결과 저장
    formatted_file = os.path.join(output_dir, '통학구역별_학생수_최종결과_개선.xlsx')
    wb.save(formatted_file)
    print(f"\n포맷이 적용된 결과가 '{formatted_file}' 파일로 저장되었습니다.")
    
    return formatted_file
    
# 개선된 메인 함수
def process_student_data_improved():
    try:
        # 파일 경로 설정
        script_dir = os.path.dirname(os.path.abspath(__file__))
        student_file = os.path.join(script_dir, '지번주소_변환결과.xlsx')
        processed_district_file = os.path.join(script_dir, '관할구역_전처리.xlsx')
        original_district_file = os.path.join(script_dir, '관할구역.xlsx')
        
        # 파일 존재 여부 확인 및 불러오기
        print(f"파일 존재 여부 확인 중...")
        for file_path in [student_file, processed_district_file, original_district_file]:
            exists = os.path.exists(file_path)
            print(f"  - {os.path.basename(file_path)}: {'있음' if exists else '없음'}")
            if not exists:
                print(f"오류: {file_path} 파일을 찾을 수 없습니다.")
                return None
        
        # 데이터 로드
        print("\n데이터 로드 중...")
        students_df = pd.read_excel(student_file)
        processed_districts_df = pd.read_excel(processed_district_file)
        original_districts_df = pd.read_excel(original_district_file)
        
        total_students = len(students_df)
        print(f"  - 학생 데이터: {total_students}개 행 로드됨")
        print(f"  - 전처리된 관할구역 데이터: {len(processed_districts_df)}개 행 로드됨")
        print(f"  - 원본 관할구역 데이터: {len(original_districts_df)}개 행 로드됨")
        
        # 1. 관할구역 데이터 심층 분석
        print("\n관할구역 데이터 분석 중...")
        # 관할구역 데이터 전처리
        district_info = preprocess_district_data(original_districts_df)
        print(f"  - {len(district_info)}개의 관할구역 정보 추출됨")
        
        # 번지 범위 추출
        district_ranges_df = parse_district_ranges(original_districts_df)
        print(f"  - {len(district_ranges_df)}개의 번지 범위 추출됨")
        
        # 동별 관할구역 요약
        district_summary = summarize_districts_by_dong(original_districts_df)
        print(f"  - {len(district_summary)}개 동의 관할구역 정보 요약됨")
        
        # 2. 지번주소 매핑 테이블 생성
        print("\n지번주소 매핑 테이블 생성 중...")
        jibun_mapping_table = create_jibun_mapping_table(district_info)
        print(f"  - {len(jibun_mapping_table)}개의 지번주소 매핑 생성됨")
        
        # 야당동 특화 매핑 테이블 추가
        yadang_mapping = create_yadang_mapping_table()
        jibun_mapping_table.update(yadang_mapping)
        print(f"  - 야당동 매핑 추가 후 총 {len(jibun_mapping_table)}개의 지번주소 매핑")
        
        # 3. 학생 데이터 전처리
        print("\n학생 데이터 전처리 중...")
        # 학년 정보 추출
        students_df['grade_num'] = students_df['학년'].apply(extract_grade_number)
        
        # 주소 성분 추출
        address_components = students_df['지번주소'].apply(extract_address_components)
        students_df['dong_name'] = [comp[0] for comp in address_components]
        students_df['main_number'] = [comp[1] for comp in address_components]
        students_df['sub_number'] = [comp[2] for comp in address_components]
        
        # 건물명 추출
        students_df['building_name'] = students_df['지번주소'].apply(extract_building_name)
        
        # 주소 표준화
        students_df['std_address'] = students_df['지번주소'].apply(standardize_address)
        
        # 4. 주소 매핑 분석
        print("\n주소 매핑 분석 중...")
        district_coverage, dong_main_counts = analyze_district_mapping_coverage(students_df, district_ranges_df)
        
        # 동별 매핑 커버리지 표시
        print("\n동별 번지 매핑 커버리지:")
        for dong, coverage in district_coverage.items():
            coverage_pct = coverage['매핑 비율'] * 100
            print(f"  - {dong}: {coverage['매핑된 번지 수']}/{coverage['총 번지 수']} 번지 매핑됨 ({coverage_pct:.1f}%)")
        
        # 5. 매핑 테이블 생성
        print("\n매핑 테이블 생성 중...")
        # processed_districts_df에서 매핑 테이블 생성
        processed_districts_df['통_cleaned'] = processed_districts_df['통'].apply(clean_tong_name)
        
        district_address_map = {}
        for _, district in processed_districts_df.iterrows():
            if pd.notna(district['지번주소']):
                district_address_map[district['지번주소']] = (district['동'], district['통_cleaned'])
        
        district_building_map = {}
        for _, district in processed_districts_df.iterrows():
            if pd.notna(district['건물명']):
                district_building_map[district['건물명']] = (district['동'], district['통_cleaned'])
        
        # 수동 매핑 테이블 업데이트
        manual_mapping.update(jibun_mapping_table)
        
# 6. 주소 매핑 적용
        print("\n학생 주소 매핑 중...")
        students_df[['mapped_dong', 'mapped_tong']] = students_df.apply(
            lambda row: pd.Series(map_address_to_school_area(row, manual_mapping, 
                                                           district_building_map, 
                                                           district_address_map, 
                                                           district_ranges_df)),
            axis=1
        )
        
        # 샘플 매핑 결과 확인
        print("\n매핑 결과 샘플:")
        sample_df = students_df.sample(min(5, len(students_df)))
        for _, row in sample_df.iterrows():
            print(f"  - {row['지번주소']} -> {row['mapped_dong']}, {row['mapped_tong']}")
        
        # 매핑 결과 분석
        mapped_count = students_df[students_df['mapped_tong'] != '미확인'].shape[0]
        unmapped_count = students_df[students_df['mapped_tong'] == '미확인'].shape[0]
        
        print(f"\n매핑 결과 요약:")
        print(f"  - 전체 학생 수: {total_students}명")
        print(f"  - 통 식별 학생 수: {mapped_count}명 ({mapped_count/total_students*100:.1f}%)")
        print(f"  - 통 미확인 학생 수: {unmapped_count}명 ({unmapped_count/total_students*100:.1f}%)")
        
        # 동-통별 분포 분석
        dong_tong_counts = students_df.groupby(['mapped_dong', 'mapped_tong']).size().reset_index(name='학생수')
        print("\n동-통별 학생 분포 (상위 10개):")
        top_areas = dong_tong_counts.sort_values('학생수', ascending=False).head(10)
        for _, row in top_areas.iterrows():
            print(f"  - {row['mapped_dong']}, {row['mapped_tong']}: {row['학생수']}명")
        
        # 7. 학년 조정 및 집계
        print("\n학년별 집계 중...")
        students_df['adjusted_grade'] = students_df['grade_num']
        non_valid_grade_mask = ~students_df['grade_num'].between(1, 6)
        students_df.loc[non_valid_grade_mask, 'adjusted_grade'] = 0  # 0은 '기타학년'을 의미
        
        # 동-통별, 학년별 학생 수 집계
        pivot_result = pd.pivot_table(
            students_df,
            values='성명',
            index=['mapped_dong', 'mapped_tong'],
            columns='adjusted_grade',
            aggfunc='count',
            fill_value=0
        )
        
        # 학년 컬럼 처리
        grade_columns = list(range(1, 7))
        for grade in grade_columns:
            if grade not in pivot_result.columns:
                pivot_result[grade] = 0
        
        # '기타학년' 컬럼 추가
        if 0 in pivot_result.columns:
            pivot_result = pivot_result.rename(columns={0: '기타학년'})
        else:
            pivot_result['기타학년'] = 0
        
        # 컬럼 순서 정리
        column_order = grade_columns + ['기타학년']
        pivot_result = pivot_result.reindex(columns=column_order)
        
        # 전체 합계 계산
        pivot_result['계'] = pivot_result.sum(axis=1)
        
        # 필요한 동-통 구역 목록
        required_areas = [
            ('운정3동', '18-19통'),
            ('운정3동', '29통'),
            ('운정3동', '30통'),
            ('운정3동', '31통'),
            ('운정4동', '8통'),
            ('운정4동', '9통'),
            ('운정4동', '10통'),
            ('운정4동', '11통'),
            ('운정4동', '12통'),
            ('운정4동', '13통'),
            ('운정4동', '15통'),
            ('운정4동', '16통'),
            ('운정4동', '17통'),
            ('운정4동', '18통'),
            ('운정4동', '10-12통'),
            ('운정3동', '미확인'),
            ('운정4동', '미확인'),
            ('파주시', '미확인'),
            ('기타', '통학구역외')
        ]
        
        # 모든 실제 매핑 결과 및 필요한 영역 포함
        final_result = []
        
        # 1. 모든 실제 매핑 결과 포함
        for index, row in pivot_result.iterrows():
            dong, tong = index
            
            row_data = {'동': dong, '통': tong}
            # 일반 학년 (1-6학년)
            for grade in range(1, 7):
                row_data[f'{grade}학년'] = int(row.get(grade, 0))
            # 기타 학년
            row_data['기타학년'] = int(row.get('기타학년', 0))
            # 합계
            row_data['계'] = int(row['계'])
            
            final_result.append(row_data)
        
        # 2. required_areas에 있지만 매핑 결과에 없는 케이스를 0으로 채움
        for dong, tong in required_areas:
            if not any(r['동'] == dong and r['통'] == tong for r in final_result):
                row_data = {'동': dong, '통': tong}
                for grade in range(1, 7):
                    row_data[f'{grade}학년'] = 0
                row_data['기타학년'] = 0
                row_data['계'] = 0
                final_result.append(row_data)
        
        result_df = pd.DataFrame(final_result)
        
        # 8. 결과 저장
        print("\n결과 파일 생성 중...")
        # 매핑 결과를 원본 파일에 추가하여 저장
        students_with_mapping = students_df.copy()
        students_with_mapping = students_with_mapping.rename(columns={
            'mapped_dong': '동',
            'mapped_tong': '통'
        })
        
        original_cols = [col for col in students_df.columns if col not in ['mapped_dong', 'mapped_tong']]
        mapping_cols = ['동', '통']
        
        mapping_result_file = os.path.join(script_dir, '지번주소_매핑결과_개선.xlsx')
        students_with_mapping[original_cols + mapping_cols].to_excel(mapping_result_file, index=False)
        print(f"  - 매핑 결과가 포함된 파일: '{mapping_result_file}'")
        
        # 기본 결과 저장
        result_file = os.path.join(script_dir, '통학구역별_학생수_결과_개선.xlsx')
        result_df.to_excel(result_file, index=False)
        print(f"  - 집계 결과 파일: '{result_file}'")
        
        # 포맷팅된 결과 생성
        formatted_file = create_formatted_result(result_df, script_dir)
        print(f"  - 포맷된 결과 파일: '{formatted_file}'")
        
        # 9. 통계 요약
        print("\n===== 통계 요약 =====")
        valid_students_count = students_df['grade_num'].between(1, 6).sum()
        other_students_count = total_students - valid_students_count
        area_totals = result_df['계'].sum()
        
        print(f"  - 총 학생 수: {total_students}명")
        print(f"  - 1-6학년 학생 수: {valid_students_count}명")
        print(f"  - 기타 학년 학생 수: {other_students_count}명")
        print(f"  - 집계된 동-통별 총 학생 수: {area_totals}명")
        
        if total_students != area_totals:
            print(f"\n주의: 집계 결과({area_totals}명)와 원본 데이터({total_students}명)의 학생 수가 일치하지 않습니다.")
            print(f"  - 차이: {total_students - area_totals}명")
        else:
            print("\n모든 학생이 정확하게 집계되었습니다!")
        
        print(f"\n*** 처리가 완료되었습니다 ***")
        
        return formatted_file
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    result_file = process_student_data_improved()
    if result_file:
        print(f"\n처리가 완료되었습니다. 결과 파일: {result_file}")
    else:
        print("\n처리 중 오류가 발생했습니다.")
