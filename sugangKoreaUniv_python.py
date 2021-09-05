from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import time
from bs4 import BeautifulSoup
from openpyxl import Workbook
import re
import pyautogui
import os
from selenium import webdriver
import sys, os

if  getattr(sys, 'frozen', False): 
    chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
    driver = webdriver.Chrome(chromedriver_path)
else:
    driver = webdriver.Chrome()
time.sleep(3)    
delay = int(input('딜레이 시간(평소 3초 권장) (입력 후 엔터) : '))
id = input('학번을 입력해주세요 (입력 후 엔터) : ')
pwd = input('비밀전호를 입력해주세요 (입력 후 엔터) : ')
num1 = int(input('전공 : 0 / 학문의 기초 : 1 (입력 후 엔터) : '))
if (num1 == 0):
    num2 = int(input('간호대학 : 0 / 경영대학 : 1 / 공과대학 : 2 / 국제학부 : 3\n 디자인조형학부 : 4 / 문과대학 : 5 / 미디어학부 : 6 / 보건과학대학 : 7 \n사범대학 : 8 / 생명과학대학 : 9 / 스마트보안학부 : 10 / 심리학부 : 11 \n의과대학 : 12 / 이과대학 : 13 / 정경대학 : 14 / 정보대학 : 15\n : '))
    if (num2 == 0):
        num3 = int(input('간호학과 : 0\n : '))
    elif (num2 == 1):
        num3 = int(input('경영학과 : 0\n : '))
    elif (num2 == 2):
        num3 = int(input('공과대학 : 0 / 기계공학부 : 1 / 산업경영공학부 : 2 / 신소재공학부 : 3\n 전지전자공학부 : 4 / 건축사회환경공학부 : 5 / 건축학과 : 6 / 기술창업융합전공 : 7 \n반도체공학과 : 8 / 융합에너지공학부 : 9 / 화학생명공학과 : 10\n :'))
    elif (num2 == 3):
        num3 = int(input('국제학부 : 0 / GKS 융합전공 : 1\n : '))
    elif (num2 == 4):
        num3 = int(input('디자인조형학부 : 0\n : '))
    elif (num2 == 5):
        num3 = int(input('문과대학 : 0 / EML 융합전공 : 1 / GLEAC 융합전공 : 2 / LB&C 융합전공 : 3\n 과학기술융합전공 : 4 / 국어국문학과 : 5 / 노어노문학과 : 6 / 독어독문학과 : 7 \n불어불문학과 : 8 / 사학과 : 9 / 사회학과 : 10 / 서어서문학과 : 11 \n         언어학과 : 12 / 영어영문학과 : 13 / 의료인문학융합전공 : 14 / 인문학과문화산업융합전공 : 15\n         인문학과정의융합전공 : 16 / 일어일문학과 : 17 / 중어중문학과 : 18 / 철학과 : 19\n         통일과국제평화융합전공 : 20 / 한국사학과 : 21 / 한문학과 : 22\n: '))
    elif (num2 == 6):
        num3 = int(input(
            '미디어학부 : 0\n : '))
    elif (num2 == 7):
        num3 = int(input(
            '물리치료학과 : 0 / 바이오시스템의과학부 : 1 / 바이오의공학부 : 2 / 보건정책관리학부 : 3\n 보건정책융합과학부 : 4\n : '))
    
    elif (num2 == 8):
        num3 = int(input(
            '가정교육과 : 0 / 교육학과 : 1 / 국어교육학과 : 2 / 다문화한국어융합전공 : 3\n 수학교육과 : 4 / 역사교육과 : 5 / 영어교육과 : 6 / 지리교육과 : 7 \n체육학과 : 8 / 패션디자인및머천다이징융합전공 : 9\n : '))
    elif (num2 == 9):
        num3 = int(input(
            '생명공학부 : 0 / 생명과학대학 : 1 / 생명과학부 : 2 / 환경생태공학부 : 3\n 기후변화융합전공 : 4 / 생태조경융합전공 : 5 / 식품공학과 : 6 / 식품자원경제학과 : 7 \n의과학융합전공 : 8\n : '))
    elif (num2 == 10):
        num3 = int(input('사이버국방학과 : 0 / 스마트보안학부 : 1\n : '))
    elif (num2 == 11):
        num3 = int(input('의예과 : 0 / 의학과 : 1\n : '))
    elif (num2 == 12):
        num3 = int(input('심리학부 : 0\n : '))
    elif (num2 == 13):
        num3 = int(input(
            '이과대학 : 0 / 물리학과 : 1 / 수학과 : 2 / 지구환경과 : 3 / 화학과 : 4\n : '))
    elif (num2 == 14):
        num3 = int(input('경제학과 : 0 / 금융공학융합전공 : 1 / 정치외교학과 : 2 / 통계학과 : 3 / 행정학과 : 4\n : '))
    elif (num2 == 15):
        num3 = int(input(
            '뇌인지과학융합전공 : 0 / 소프트웨어벤쳐융합전공 : 1 / 인공지능융합전공 : 2 / 정보보호융합전공 : 3 / 컴퓨터학과 : 4\n : '))
    else:
        print('잘못된 정보')
elif (num1 == 1):
    num2 = int(input('경영대학 : 0 / 공과대학 : 1 / 디자인조형학부 : 2 / 문과대학 : 3\n보건과학대학 : 4 / 생명과학대학 : 5 \n스마트보안학부 : 6 / 정경대학 : 7 / 정보대학 : 8\n : '))
    if (num2 == 0):
        num3 = int(input('경영학과 : 0\n : '))
    elif (num2 == 1):
        num3 = int(input('공과대학 : 0 / 융합에너지공학과 : 1 \n : '))
    elif (num2 == 2):
        num3 = int(input('디자인조형학부 : 0 \n : '))
    elif (num2 == 3):
        num3 = int(input('노어노문학과 : 0 / 독어독문학과 : 1 / 불어불문학과 : 2 / 서어서문학과 : 3 \n         영어영문학과 : 4 / 일어일문학과 : 5 / 중어중문학과 : 6\n : '))
    elif (num2 == 4):
        num3 = int(input('보건정책관리학부 : 0\n : '))
    elif (num2 == 5):
        num3 = int(input('식품자원경제학과 : 0\n : '))
    elif (num2 == 6):
        num3 = int(input('스마트보안학부 : 0\n : '))    
    elif (num2 == 7):
        num3 = int(input('통계학과 : 0 / 행정학과 : 1\n : '))
    elif (num2 == 8):
        num3 = int(input('컴퓨터학과 : 0\n : '))
    else:
        print('잘못된 정보')
else:
    print('잘못된 정보')

URL = 'https://sugang.korea.ac.kr/'
driver.get(URL)

html = driver.page_source
time.sleep(delay)

# 경고메세지 없애기
driver.switch_to_frame('Main')
driver.find_element_by_css_selector('div.jconfirm-closeIcon').click()

# 로그인
elem = driver.find_element_by_id("id")
elem.send_keys(id)
elem = driver.find_element_by_id("pwd")
elem.send_keys(pwd)
driver.find_element_by_xpath('//*[@id="btn-login"]').click()
time.sleep(delay)

# 과목조회
driver.switch_to_frame('coreMain')
driver.find_element_by_xpath('//*[@id="menu_hakbu"]').click()
time.sleep(delay)

# 이수구분
sel1 = [driver.find_elements_by_xpath('//*[@id="pCourDiv"]/option')[0], driver.find_elements_by_xpath('//*[@id="pCourDiv"]/option')[1]]
s1 = sel1[num1]
s1.click()
time.sleep(1)
sel2 = driver.find_elements_by_xpath('//*[@id="pCol"]/option')

s2 = sel2[num2]
s2.click()
time.sleep(1)

# 엑셀 생성
wb = Workbook()
sel3 = driver.find_elements_by_xpath('//*[@id="pDept"]/option')

s3 = sel3[num3]
s3.click()
time.sleep(delay)
major_name = s3.text  # 전공 이름
wb.create_sheet(title=major_name)

driver.find_element_by_xpath('//*[@id="btnSearch"]').click()  # 조회버튼
time.sleep(delay)
soup = BeautifulSoup(driver.page_source, 'lxml')
tr = soup.select(
'html body div div div div div div table tbody tr')
ws = wb[major_name]
ws.append(['캠퍼스', '학수번호', '분반', '이수구분', '교과목명', '담당교수', '1학년(신청/정원)', '2학년(신청/정원)','3학년(신청/정원)', '4학년(신청/정원)', '교환학생(신청/정원)', '대학원(신청/정원)', '전체(신청/정원)'])
for tr_info in tr:
    lst = []
    for idx, inf in enumerate(tr_info):
        if (idx % 15 == 1 or idx % 15 == 2 or idx % 15 == 3 or idx % 15 == 4 or idx % 15 == 5 or idx % 15 == 6):
            p = re.compile(r'<.+?>')
            inf = re.sub(p, '', str(inf))
            lst.append(inf)

    # 희망과목 정보
    p = re.compile('fnAplyState.*\)')
    result = re.search(p, str(tr_info))
    if (result == None ):
        continue
    script = result.group(0)
    driver.execute_script(script)
    time.sleep(delay)
    soup = BeautifulSoup(driver.page_source, 'lxml')
    # 1,2,3,4,교환,대학원,전체
    yr = soup.select('html body div div div div div div div div div div div div div table tbody tr th')
    num = soup.select('html body div div div div div div div div div div div div div table tbody tr td')
    p = re.compile(r'<.+?>')
    yr = re.sub(p, '', str(yr))
    # 희망과목 정보 전처리
    ap_list = []
    t_list = []
    result = []
    for idx, n in enumerate(num):
        p = re.compile(r'<.+?>')
        n_info = re.sub(p, '', str(n))
        if(n_info == ''):
            n_info = '0'
        if(idx % 2 == 0):
            ap_list.append(n_info)
        else:
            t_list.append(n_info)

    for i in range(7):
        result.append((ap_list[i]+'/'+t_list[i]))
    ws.append(lst+result)
    elem = driver.find_element_by_xpath('/html/body/div[3]/div[2]/div/div/div/div/div/div/div/div[1]')
    elem.click()
total_name = s1.text + '_' + s2.text + '_' + s3.text
wb.remove(wb['Sheet'])
wb.save(total_name + '.xlsx')
wb.close



