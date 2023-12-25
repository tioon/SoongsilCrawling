import requests
from lxml import html
import pandas as pd
from datetime import datetime

def GetCompetitionRate(SchoolName, TagName, UrlName): 
    response = requests.get(UrlName,verify=False)
    rate_data = None
    
    if response.status_code == 200:
        tree = html.fromstring(response.content)
        CompetitionRate = tree.xpath(TagName+'/text()')
        
        if CompetitionRate:
            rate_data = {"학교": SchoolName, "논술 경쟁률": CompetitionRate[0]}
        else: 
            print(f"HTML코드에서 {TagName} 클래스를 가진 태그를 찾을 수 없습니다.")
    else:
        print(f"웹 페이지 요청 오류: {response.status_code}")

    return rate_data

# 논술 경쟁률 페이지 정보
schools = [
    ("숭실대학교", "//td[@class=\"rate1\"]", "https://addon.jinhakapply.com/RatioV1/RatioH/Ratio11010301.html"),
    ("건국대학교", "//td[@class=\"rate1\"]", "https://addon.jinhakapply.com/RatioV1/RatioH/Ratio10080151.html"),
    ("동국대학교", "//td[@class=\"rate1\"]", "https://addon.jinhakapply.com/RatioV1/RatioH/Ratio10550291.html"),
    ("홍익대학교", "//td[@class=\"rate1\"]", "https://addon.jinhakapply.com/RatioV1/RatioH/Ratio11720351.html"),
    ("세종대학교", "//td[@class=\"rate1\"]", "https://addon.jinhakapply.com/RatioV1/RatioH/Ratio10950471.html"),
    ("아주대학교", "//td[@class=\"rate1\"]", "https://addon.jinhakapply.com/RatioV1/RatioH/Ratio11040371.html")
]

data_list = []

for school in schools:
    data = GetCompetitionRate(*school)
    if data:
        data_list.append(data)

df = pd.DataFrame(data_list)

# 현재의 날짜와 시간을 가져와서 파일 이름에 추가
current_time = datetime.now().strftime('%Y-%m-%d %H시 %M분')
filename = f"학교 논술 경쟁률_{current_time}.xlsx"
df.to_excel(filename, index=False)

with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name="Sheet1")

    # 엑셀 열의 너비 조절
    worksheet = writer.sheets["Sheet1"]
    worksheet.column_dimensions['A'].width = 20  # 1열의 너비를 20으로 설정
    worksheet.column_dimensions['B'].width = 20  # 2열의 너비를 20으로 설정
