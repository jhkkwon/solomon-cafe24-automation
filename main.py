import xlwt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import re
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()
ID = os.environ.get("ID")
PASSWORD = os.environ.get("PASSWORD")
# 전체 데이타
data = []
# 각 행 데이터(고객별)
# data 리스트에 append 해줌,


options = webdriver.ChromeOptions()

# headless 옵션 설정
options.add_argument('headless')
options.add_argument("no-sandbox")

# 브라우저 윈도우 사이즈
options.add_argument('window-size=1920x1080')

# 사람처럼 보이게 하는 옵션들
options.add_argument("disable-gpu")  # 가속 사용 x
options.add_argument("lang=ko_KR")  # 가짜 플러그인 탑재
options.add_argument(
    'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36')  # user-agent 이름 설정

# 드라이버 위치 경로 입력
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# url을 이용하여 브라우저로 접속
driver.get('https://eclogin.cafe24.com/Shop/')

driver.implicitly_wait(3)

# 아이디/비밀번호를 입력하기
driver.find_element(by=By.XPATH,
                    value='/html/body/div[2]/div/section/div/form/div/div[1]/div/div[1]/div/input').send_keys(
    ID)
driver.find_element(by=By.XPATH,
                    value='/html/body/div[2]/div/section/div/form/div/div[1]/div/div[2]/div/input').send_keys(
    PASSWORD)

# 로그인 버튼을 누르기
driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/section/div/form/div/div[3]/button').click()

# 비밀번호 변경 페이지 (아니요 누르기)
driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/div/div[1]/div[4]/a[2]').click()

# 팝업페이지 x 클릭
driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/div/button').click()

# 주문 관리 누르기
driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[1]/div[2]/ul/li[3]/a').click()

# 배송 준비중 관리 누르기
driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[2]/div[1]/div/div[2]/ul/li[4]/a').click()

# 검색 누르기
driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[2]/div[2]/form[2]/div[2]/div[5]/a[1]/span').click()

# 배송준비중 목록
table = driver.find_elements(by=By.XPATH, value='/html/body/div[1]/div[2]/div[2]/form[2]/div[9]/div[6]/table/tbody')
print(table)

# 목록들을 저장함
for i in range(len(table)):
    # 멤버 임시저장할 리스트
    tmp_data = []
    # 아이템/ 옵션 / 수량 임시 저장 딕셔너리
    item_dic = {}
    # 옵션/ 수량 리스트
    opt_qty = []
    # 한 고객의 앞 주소
    MEMBER_PREFIX = '/html/body/div[1]/div[2]/div[2]/form[2]/div[9]/div[6]/table/tbody['

    cnt = 1

    # 고객 별로 정보 수집
    item_name = driver.find_element(by=By.XPATH,
                                    value=MEMBER_PREFIX + str(
                                        i + 1) + ']' + '/tr[' + str(cnt) + ']/td[10]/div/p[1]').text
    item_name = item_name.replace('마켓상품', '').replace('내역', '')
    item_name = re.sub(pattern='\([^)]*\)', repl='', string=item_name)
    item_name = item_name.strip()
    print(item_name)

    try:
        option = driver.find_element(by=By.XPATH,
                                     value=MEMBER_PREFIX + str(
                                         i + 1) + ']' + '/tr[' + str(cnt) + ']/td[10]/div/ul/li').text
        print(option)

    except:
        option = ''
        pass

    quantity = driver.find_element(by=By.XPATH,
                                   value=MEMBER_PREFIX + str(
                                       i + 1) + ']' + '/tr[' + str(cnt) + ']/td[12]').text
    print(quantity)

    item_dic[item_name] = option + ' * ' + quantity
    cnt += 1

    try:
        while (True):
            item_name = driver.find_element(by=By.XPATH, value=MEMBER_PREFIX + str(i + 1) + ']' + '/tr[' + str(
                cnt) + ']/td[3]/div/p[1]').text
            item_name = item_name.replace('마켓상품', '').replace('내역', '')
            item_name = re.sub(pattern='\([^)]*\)', repl='', string=item_name)
            print(item_name.strip())

            try:
                option = driver.find_element(by=By.XPATH, value=MEMBER_PREFIX + str(i + 1) + ']' + '/tr[' + str(
                    cnt) + ']/td[3]/div/ul/li').text
                print(option)

            except:
                option = ''
                pass

            quantity = driver.find_element(by=By.XPATH,
                                           value=MEMBER_PREFIX + str(i + 1) + ']' + '/tr[' + str(cnt) + ']/td[5]').text

            print(quantity)

            if item_name in item_dic:
                item_dic[item_name] = item_dic[item_name] + ' , ' + option + ' * ' + quantity
            else:
                item_dic[item_name] = option + ' * ' + quantity
            cnt += 1
    except:
        name = driver.find_element(by=By.XPATH,
                                   value=MEMBER_PREFIX + str(i + 1) + ']/tr[' + str(cnt) + ']/td/strong/span').text
        print(name)

        address_li = driver.find_element(by=By.XPATH,
                                         value=MEMBER_PREFIX + str(i + 1) + ']/tr[' + str(cnt) + ']/td/ul/li[3]').text
        address = ' '.join(address_li.split()[3:])
        print(address)

        phone_num_1_li = driver.find_element(by=By.XPATH, value=MEMBER_PREFIX + str(i + 1) + ']/tr[' + str(
            cnt) + ']/td/ul/li[1]').text
        phone_num_1 = ' '.join(phone_num_1_li.split()[2:])
        print(phone_num_1)

        phone_num_2_li = driver.find_element(by=By.XPATH, value=MEMBER_PREFIX + str(i + 1) + ']/tr[' + str(
            cnt) + ']/td/ul/li[2]').text
        phone_num_2 = ' '.join(phone_num_2_li.split()[2:])
        print(phone_num_2)

        root = driver.find_element(by=By.XPATH, value=MEMBER_PREFIX + str(i + 1) + ']/tr[1]/td[3]/img').get_attribute(
            'title')
        print(root)

        try:
            driver.find_element(by=By.XPATH, value=MEMBER_PREFIX + str(i + 1) + ']/tr[1]/td[3]/a/span').click()
            driver.switch_to.window(driver.window_handles[-1])
            message = driver.find_element(by=By.XPATH, value='//*[@id="DeliveryMemo"]').text
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            print(message)

        except:
            message = ''
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        print('***' * 10)

    tmp_data.append(name)
    tmp_data.append(address)
    tmp_data.append(phone_num_1)
    tmp_data.append(phone_num_2)
    tmp_data.append(message)
    tmp_data.append(' //'.join(list(item_dic)))
    tmp_data.append(' // '.join(list(item_dic.values())))
    tmp_data.append('1')
    tmp_data.append('2750')
    tmp_data.append('030')
    tmp_data.append(root)

    data.append(tmp_data)

wb = xlwt.Workbook(encoding='utf-8')
ws = wb.add_sheet('sheet1')

# 칼럼 타이틀 설정
title = ['수령인', '수령인 주소', '수령인 전화번호', '수령인 휴대전화', '배송메세지', '주문상품명', '상품옵션', '수량', '총 배송비', '상품배송유형', '주문 경로']
for kwd, j in zip(title, list(range(len(title)))):
    ws.write(0, j, kwd)

ws.col(0).width = 12 * 255
ws.col(1).width = 57 * 255
ws.col(2).width = 17 * 255
ws.col(3).width = 17 * 255
ws.col(4).width = 25 * 255
ws.col(5).width = 56 * 255
ws.col(6).width = 30 * 255

for row_num in range(1, 1 + len(data)):
    for col_num in range(len(title)):
        ws.write(row_num, col_num, data[row_num - 1][col_num])

now = datetime.now()
time_now = now.strftime('%Y%m%d_%H%M%S')
wb.save('./로젠 택배_' + time_now + '.xls')
