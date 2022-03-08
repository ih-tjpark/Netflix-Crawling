import smtplib
from email.mime.multipart import MIMEMultipart
from email import encoders
import pandas as pd
import os
import time

# 워드 클라우드
import konlpy
from konlpy.tag import Okt
from collections import Counter
import numpy as np
from PIL import Image
from wordcloud import WordCloud
import matplotlib.pyplot as plt

from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

# 텍스트 형식
from email.mime.text import MIMEText
# 이미지 형식
from email.mime.image import MIMEImage
# 오디오 형식
from email.mime.audio import MIMEAudio

# 위의 모든 객체들을 생성할 수 있는 기본 객체
# MIMEBase(_maintype, _subtype)
# MIMEBase(<메인 타입>, <서브 타입>)
from email.mime.base import MIMEBase

# 이미지 첨부를 위한 Class
class imageAttach(object):
    def __init__(self,name):
        self.name = name

    def __str__(self):
        return self.name

    def __repr__(self):
        return "'"+self.name+"'"

# 메일 전송 함수
def send_email(smtp_info, msg):
    with smtplib.SMTP(smtp_info["smtp_server"], smtp_info["smtp_port"]) as server:
        # TLS 보안 연결
        server.starttls()
        # 로그인
        server.login(smtp_info["smtp_user_id"], smtp_info["smtp_user_pw"])

        # 로그인 된 서버에 이메일 전송
        response = server.sendmail(msg['from'], msg['to'], msg.as_string())
        # 메시지를 보낼때는 .as_string() 메소드를 사용해서 문자열로 변경

        # 이메일을 성공적으로 보내면 결과는 {}
        if not response:
            print('이메일을 성공적으로 보냈습니다.')
        else:
            print(response)

# 첨부 파일 추가 함수
def make_multimsg(msg_dict):
    multi = MIMEMultipart(_subtype='mixed')

    for key, value in msg_dict.items():
        # 각 타입에 적절한 MIMExxx()함수를 호출하여 msg 객체를 생성.
        if key == 'text':
            with open(value['filename'], encoding='utf-8') as fp:
                msg = MIMEText(fp.read(), _subtype=value['subtype'])
        elif key == 'image':
            with open('./result/'+value['filename'], 'rb') as fp:
                msg = MIMEImage(fp.read(), _subtype=value['subtype'])
                print("첨부완료")
        elif key == 'audio':
            with open(value['filename'], 'rb') as fp:
                msg = MIMEAudio(fp.read(), _subtype=value['subtype'])
        else:
            with open('./result/'+value['filename'], 'rb') as fp:
                msg = MIMEBase(value['maintype'],  _subtype=value['subtype'])
                msg.set_payload(fp.read())
                encoders.encode_base64(msg)
        # 파일 이름을 첨부파일 제목으로 추가
        msg.add_header('Content-Disposition', 'attachment', filename=value['filename'])

        # 첨부파일 추가
        multi.attach(msg)

    return multi

# 웹 크롤링 함수
def selenium_Crawling(read_title):

    start_url = 'http://www.naver.com' # 첫 시작 url

    driver.get(start_url)
    search = driver.find_element(By.ID,'query') #검색창 가져오기
    time.sleep(1)
    search.send_keys(f"넷플릭스 {read_title} 리뷰") #검색창 입력

    driver.find_element(By.ID,'search_btn').click() #버튼 클릭
    time.sleep(1)

    driver.find_element(By.LINK_TEXT,"VIEW").click() #view 클릭
    time.sleep(1)

    driver.find_element(By.LINK_TEXT,"블로그").click() #블로그 클릭
    time.sleep(1)

    des_texts = driver.find_elements(By.CLASS_NAME,'dsc_txt') #미리보기 text 가져오기

    des_text=""
    for i in des_texts:
        des_text += " "+i.text
    print(des_text)
    return des_text

# 전처리 및 워드카운트
def preprocessing(read_text,stop_words):
    # 토큰화
    okt=Okt()
    #print(des_texts)
    # 토큰화 후 배열로 저장
    line=[]
    line = okt.pos(read_text)

    n_adj=[]
    for word, tag in line:
        if tag in ['Noun','Adjective']:
            n_adj.append(word)

    # 불용어 제외한 단어만 남기기
    n_adj = [word for word in n_adj if not word in stop_words]

    # 가장 많이 나온 단어 50개 저장
    counts = Counter(n_adj)
    tags = counts.most_common(50)
    return tags

# 워드클라우드 생성 함수
def toWordcloud(tags,mask,font):
    # 워드클라우드 generate
    word_cloud = WordCloud(font_path=font,
                           background_color='black',
                           max_font_size=300,
                           mask=mask, colormap='Reds').\
        generate_from_frequencies(dict(tags))

    # 그림으로 표현하기
    plt.figure(figsize=(20,8))
    plt.title(titles)
    plt.imshow(word_cloud)
    plt.axis('off')

    # 이미지로 저장
    word_cloud.to_file(f'./result/{titles}.png')


# Netflix 나라별 Top10 엑셀에서 최신 한국기준으로 불러오기
excel_df = pd.read_excel('all-weeks-countries.xlsx',
                         usecols = ['country_name','week','category','show_title'])
excel_df = excel_df[excel_df['country_name']=='South Korea']
excel_df = excel_df[excel_df['week']=='2022-02-27']


# 제목리스트 추출
title_list = excel_df.show_title.tolist()

print(title_list)

# Web Crawling 셋팅
chrome_options = ChromeOptions()
prefs = {"download.default_directory": os.getcwd()}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument('--kiosk-printing')
chrome_options.add_argument('start-maximized')
chrome_options.add_argument("disable-gpu")
driver = Chrome(ChromeDriverManager().install(), options=chrome_options)

maximum_wait = 15
wait = WebDriverWait(driver, maximum_wait)


# 불용어 파일 셋팅
stop_words = open("koreanStopwords.txt",'r',encoding='utf8').read()
stop_words=set(stop_words.split('\n'))

#wordcloud 설정
mask = Image.new("RGBA",(920,920),(255,255,255))
image=Image.open('netflix_logo.png').convert("RGBA")
x,y =image.size
mask.paste(image,(0,0,x,y),image)
mask = np.array(mask)
font ='t0SbEaWPJanI7N31lVxO_i1Ngp0.ttf' # 폰트


attach_dict={}
tag_list=[]

# 타이틀 별 웹 크롤링, 전처리, 시각화 진행
for titles in title_list:

    # 웹 크롤링 실행
    des_text = selenium_Crawling(titles)

    # 형태소 분석 & 워드 카운트 실행
    tags = preprocessing(des_text,stop_words)

    # 워드 클라우드 실행
    toWordcloud(tags,mask,font)

    # 타이틀 별 카운트된 워드 Top5 리스트에 저장
    top5_tag=tags[0][0]
    for num in range(1,5):
        top5_tag += ', '+tags[num][0]
    print(top5_tag)
    tag_list.append(top5_tag)

    # 이메일 전송용 이미지데이터 저장
    attach_dict[imageAttach("image")] = \
        {'maintype' : 'image', 'subtype' :'png', 'filename' : titles+'.png'}

    print("=========한 작품끝 =================================================================================================================")


# 데이터프레임 엑셀로 저장
excel_df['top_word']=tag_list
excel_df.to_excel('./result/Netflix_Top10.xlsx',index=False)
attach_dict['excel'] ={'maintype' : 'excel', 'subtype' : 'octect-stream',
                       'filename' : 'Netflix_Top10.xlsx'}


# email / password 설정
email = ''
password = ''

#전송 정보 설정
smtp_info = dict({"smtp_server" : "smtp.naver.com", # SMTP 서버 주소
                  "smtp_user_id" : email,
                  "smtp_user_pw" : password ,
                  "smtp_port" : 587}) # SMTP 서버 포트

# 메일 내용 작성
title = "Netflex Top 10 영화 & TV 워드클라우드"
content = "Netflex Top 10 영화 & TV 워드클라우드 전송드립니다."
sender = "hahxowns@naver.com"
receiver = "hahxowns@gmail.com"



# 메세지 전송
msg = MIMEText(_text = content, _charset = "utf-8") # 메일 내용

# 첨부파일 추가 및 송신 메일 설정
multi = make_multimsg(attach_dict)
multi['subject'] = title    # 메일 제목
multi['from'] = sender      # 송신자
multi['to'] = receiver      # 수신자
multi.attach(msg)

 # 메일 전송
send_email(smtp_info, multi)

multi['to'] = ''      # 수신자
# 메일 전송
send_email(smtp_info, multi)