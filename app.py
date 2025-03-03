import requests

def get_city_codes():
    # 비례대표선거 시도 code 조회
    city_url = "http://info.nec.go.kr/bizcommon/selectbox/selectbox_cityCodeBySgJson.json"
    election_params = {
        'electionId': '0020240410',
        'electionCode': '7'
    }
    city_data = {}
    
    try:
        response = requests.get(city_url, params=election_params)
        data = response.json()
        
        for city in data['jsonResult']['body']:            
            city_obj = {
                'code': city['CODE'],
                'name': city['NAME'],
                'towns': get_town_codes(city['CODE'])  # 해당 시도의 구시군 정보 추가
            }
            city_data[city['CODE']] = city_obj
            
        return city_data
    
    except Exception as e:
        print(f"시도 코드 가져오기 실패: {e}")
        return None

def get_town_codes(city_code):
    # 비례대표선거 구시군 code 조회
    town_url = "http://info.nec.go.kr/bizcommon/selectbox/selectbox_townCodeJson.json"
    election_params = {
        'electionId': '0020240410',
        'electionCode': '7',
        'cityCode': city_code
    }
    
    try:
        response = requests.get(town_url, params=election_params)
        data = response.json()
        
        # 구시군 정보를 객체 배열로 반환
        return [{'code': town['CODE'], 'name': town['NAME']} for town in data['jsonResult']['body']]
        
    except Exception as e:
        print(f"구시군 코드 가져오기 실패 (시도코드: {city_code}): {e}")
        return []

def get_election_data(city_code, town_code):
    url = "http://info.nec.go.kr/electioninfo/electionInfo_report.xhtml"
    params = {
        'electionId': '0020240410',
        'requestURI': '/electioninfo/0020240410/vc/vccp08.jsp',
        'topMenuId': 'VC',
        'secondMenuId': 'VCCP08',
        'menuId': 'VCCP08',
        'statementId': 'VCCP08_#7_1',
        'electionCode': '7',
        'cityCode': city_code,
        'sgaCityCode': '-1',
        'townCodeFromSgg': '-1',
        'townCode': town_code,
        'sggTownCode': '-1',
        'checkCityCode': '-1'
    }

    response = requests.get(url, params=params)
    print(response.text)

# 실행
city_data = get_city_codes()

# for city_code, city_info in city_data.items():
#     print(f"\n{city_info['name']} : {city_info['code']}")
#     for town in city_info['towns']:
#         print(f"  - {town['name']} : {town['code']}")

get_election_data('4900', '4901')