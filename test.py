# 1. 인터파크 투어 사이트에서 여행지를 입력후 검색 -> 잠시 후 -> 결과
# 2. 로그인시 PC 웹 사이트에서 처리가 어려울 경우 -> 모바일 로그인 진입
# 3. 모듈 가져오기
from selenium import webdriver as wd
from selenium.webdriver.common.by import By
import time
from Tour import TourInfo
# 명시적 대기를 위해
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from mpmath import e1
from werkzeug.urls import _URLTuple
# Beautiful soup
from bs4 import BeautifulSoup as bs
from DbMgr import DBHelper as Db


# 4. 사전에 필요한 정보를 로드 -> 디비호고스 쉘, 배치 파일에서 인자로 받아서 세팅
db = Db()
main_url = 'http://tour.interpark.com/'
keyword = '로마'
# 상품정보를 담는 리스트 (TourInfo 리스트)
tour_list = []


# 5. 드라이버 로드
# browser = wd.Chrome('D:\CODE\python\chromedriver')
browser = wd.Chrome(executable_path='D:\CODE\python\chromedriver')
# 차후 -> 옵션 부여하여 (프록시, 에이전트 조작, 이미지를 배제)
# 크롤링을 오래돌리면 => 임시파일들이 쌓인다!! -> 임시저장파일들 삭제

# 6. 사이트 접속
browser.get(main_url)

# 7. 검색창을 찾아서 검색어 입력
browser.find_element_by_id('SearchGNBText').send_keys(keyword)
# 수정할경우 -> 뒤에 내용이 붙어버림 => .clear() -> send_kyes('내용')

# 8. 검색 버튼 클릭
browser.find_element_by_css_selector('button.search-btn').click()

# 9. 잠시 대기 => 페이지가 로드되고 나서 즉각적으로 데이터를 즉각적으로 획득하는 행위는 자제
# 명시적 대기 => 특정 요소가 로케이트(발견될때까지) 대기
try:
    element = WebDriverWait(browser, 10).until(
        # 지정한 한개 요소가 올라오면 웨이트 종료
        EC.presence_of_all_elements_located( (By.CLASS_NAME, 'oTravelBox' ) )
    )
except Exception as e:
    print( '오류발생 : ', e)
# 암묵적 대기 => DOM이 다 로드 될때까지 대기 하고 먼저 로드되면 바로 진행
# 요소를 찾을 특정 시간 동안 DOM 풀링을 지시. 10초 이내라도 발견되면 진행
browser.implicitly_wait( 10 )
# 절대기 => time.sleep(10)  -> 클라우드 페어(디도스 방어 솔루션)
# 더보기 눌러스 게시판 진입
browser.find_element_by_css_selector('.oTravelBox > .boxList > .moreBtnWrap > .moreBtn ').click()

# 게시판에서 데이터를 가져올 때
# 데이터가 많으면 세션(혹시 로그인을 해서  접근해야하는 사이트일 경우)관리
# 특정 단위별로 로그아웃 로그인 계속 시도
# 특정 게시물이 사라질 경우 -> 팝업 발생(없는 글입니다 ...) => 팝업처리검토
# 게시판  스캔시 => 임계점을 모름!! 어디가 끝인가??
# 게시판을 스캔해서 메타 정보를 획득 => loop 돌려서 일괄적으로 방문 접근 처리

# searchModule.SetCategoryList(1, '') 스크립트 실행
# 16은 임시값, 게시물을 넘어갔을때 현상을 확인차
for page in range(1,2): # 17):
    try:
        # 자바스크립트 구동하기
        browser.execute_script("searchModule.SetCategoryList(%s, '')" % page)
        time.sleep(2)
        # print('%s 페이지 이동' % page)
        #################################################################
        # 여러 사이트에서 정보를 수집할 경우 공통 정보 정의 단계 필요
        # 상품명, 코멘트, 기간1, 기간2, 가격, 평점, 섬네일, 링크(실제상품 상세정보)
        boxItems = browser.find_elements_by_css_selector(
            '.oTravelBox > .boxList > .boxItem')
        for li in boxItems:
            # 이미지를 링크값을 사용할 것인가?
            # 직접 다운로드해서 우리 서버에 업로드(ftp) 할것인가?
            # print('섬네일 : '+li.find_element_by_css_selector('img').get_attribute('src'))
            # print('링크 : '+li.find_element_by_css_selector('a').get_attribute('onclick'))
            # print('상품명 : '+li.find_element_by_css_selector('h5.proTit').text)
            # print('코멘트 : '+li.find_element_by_css_selector('.proSub').text)
            # print('가격 : '+li.find_element_by_css_selector('.proPrice').text)
            area = ''
            # for info in li.find_elements_by_css_selector('.info-row .proInfo'):
            #     print(info.text)
            # print('='*100)
            # 데이터 모음
            # li.find_elements_by_css_selector('.info-row .proInfo')[1].text,
            # 데이터가 부족하거나 없을수도 있으므로 직접 인덱스로 접근은 위험성이 있음
            obj = TourInfo(
                li.find_element_by_css_selector('h5.proTit').text,
                li.find_element_by_css_selector('.proPrice').text,
                li.find_elements_by_css_selector('.info-row .proInfo')[1].text,
                li.find_element_by_css_selector('a').get_attribute('onclick'),
                li.find_element_by_css_selector('img').get_attribute('src'))
            tour_list.append( obj )
    except Exception as e1:
        print('오류', e1)

# print(tour_list, len(tour_list))

# 수집한 정보 개수를 루프 -> 페이지 방문 -> 콘텐츠 획득(상품상세정보) => 디비
for tour in tour_list:
    # tour => TourInfo
    # print(type(tour))
    # 링크 데이터에서 실데이터 획득
    # 분해
    arr = tour.link.split(',')
    if arr:
        # 대체
        link = arr[0].replace('searchModule.OnClickDetail(','')
        # 슬라이싱 => 앞에 ', 뒤에 ' 제거
        detail_url = link[1:-1]
        # 상세페이지 이동 : URL 값이 완성된 형태인지 확인 (http://~)
        browser.get(detail_url)
        time.sleep(2)
        # pip install bs4
        # 현재 페이지를 beautifulsoup의 dom으로 구성
        soup = bs(browser.page_source, 'html.parser')
        # 현재 상세정보페이지에서 스케줄 정보를 획득
        data = soup.select('.tip-cover')
        print("--------------------")
        print(str(data[0]))
        print("--------------------")


        # content_final = ''
        # for c in data[0].contents:
        #     content_final = str(c)

        # 콘텐츠  내용에 따라 전처리 => data[0].contents
        db.db_insertCrawlingData(
            tour.title,
            tour.price,
            tour.area,
            str(data[0]),
            keyword
        )
        # 디비 입력



# 종료
browser.close()
browser.quit()
import sys
sys.exit() # 프로세스 끝내기
