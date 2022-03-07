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


class imageAttach(object):
    def __init__(self,name):
        self.name = name

    def __str__(self):
        return self.name

    def __repr__(self):
        return "'"+self.name+"'"

def send_email(smtp_info, msg):
    with smtplib.SMTP(smtp_info["smtp_server"], smtp_info["smtp_port"]) as server:
        # TLS 보안 연결
        server.starttls()
        # 로그인
        server.login(smtp_info["smtp_user_id"], smtp_info["smtp_user_pw"])

        # 로그인 된 서버에 이메일 전송
        response = server.sendmail(msg['from'], msg['to'], msg.as_string()) # 메시지를 보낼때는 .as_string() 메소드를 사용해서 문자열로 바꿔줍니다.

        # 이메일을 성공적으로 보내면 결과는 {}
        if not response:
            print('이메일을 성공적으로 보냈습니다.')
        else:
            print(response)


def make_multimsg(msg_dict):
    multi = MIMEMultipart(_subtype='mixed')

    for key, value in msg_dict.items():
        # 각 타입에 적절한 MIMExxx()함수를 호출하여 msg 객체를 생성한다.
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

excel_df = pd.read_excel('all-weeks-countries.xlsx', usecols = ['country_name','week','category','show_title'])

# 엑셀
excel_df = excel_df[excel_df['country_name']=='South Korea']
excel_df = excel_df[excel_df['week']=='2022-02-27']

# 제목만 추출
title_list = excel_df.show_title.tolist()


print(title_list)


chrome_options = ChromeOptions()

prefs = {"download.default_directory": os.getcwd()}
chrome_options.add_experimental_option("prefs", prefs)

chrome_options.add_argument(
'--kiosk-printing'
)

chrome_options.add_argument(
    'start-maximized'
)

chrome_options.add_argument(
    "disable-gpu"
)

driver = Chrome(
    ChromeDriverManager().install(), options=chrome_options
)

maximum_wait = 15
wait = WebDriverWait(driver, maximum_wait)

start_url = 'http://www.naver.com' # 첫 시작 url

value = title_list

attach_dict={}
tag_list=[]

for titles in value:
    driver.get(start_url)
    search = driver.find_element_by_id('query') #검색창 가져오기
    time.sleep(1)
    search.send_keys(f"{titles} 리뷰") #검색창 입력

    driver.find_element_by_id('search_btn').click() #버튼 클릭
    time.sleep(1)

    driver.find_element_by_link_text("VIEW").click() #view 클릭
    time.sleep(1)

    driver.find_element_by_link_text("블로그").click() #블로그 클릭
    time.sleep(1)

    des_texts = driver.find_elements_by_class_name('dsc_txt') #미리보기 text 가져오기
    des_text=""
    for i in des_texts:
        des_text += " "+i.text
    print(des_text)
    # 워드클라우드
    # 토큰화
    okt=Okt()
    #print(des_texts)
    # 토큰화 후 배열로 저장
    line=[]
    line = okt.pos(des_text)

    n_adj=[]
    for word, tag in line:
        if tag in ['Noun','Adjective']:
            n_adj.append(word)

    # 불용어 제거
    stop_words = ""
    stop_words=set(stop_words.split(' '))

    # 불용어 제외한 단어만 남기기
    n_adj = [word for word in n_adj if not word in stop_words]

    # 가장 많이 나온 단어 50개 저장
    counts = Counter(n_adj)
    tags = counts.most_common(50)


    mask = Image.new("RGBA",(920,920),(255,255,255))
    image=Image.open('netflix_logo.png').convert("RGBA")
    x,y =image.size
    mask.paste(image,(0,0,x,y),image)
    mask = np.array(mask)



    # 폰트
    font ='t0SbEaWPJanI7N31lVxO_i1Ngp0.ttf'
    # 워드클라우드 generate
    word_cloud = WordCloud(font_path=font,background_color='black',max_font_size=300, mask=mask, colormap='Reds').generate_from_frequencies(dict(tags))

    # 그림으로 표현하기
    plt.figure(figsize=(20,8))
    plt.title(titles)
    plt.imshow(word_cloud)
    plt.axis('off')

    # 이미지로 저장
    word_cloud.to_file(f'./result/{titles}.png')
    #plt.show()

    top5_tag=tags[0][0]
    for num in range(1,5):
        top5_tag += ', '+tags[num][0]
    print(top5_tag)

    tag_list.append(top5_tag)
    # 이메일 전송용 이미지데이터 저장
    attach_dict[imageAttach("image")] = {'maintype' : 'image', 'subtype' :'png', 'filename' : titles+'.png'}

    print("=========한 작품끝 =================================================================================================================")


excel_df['top_word']=tag_list
excel_df.to_excel('./result/Netflix_Top10.xlsx',index=False)
attach_dict['excel'] ={'maintype' : 'excel', 'subtype' : 'octect-stream', 'filename' : 'Netflix_Top10.xlsx'}
id =''
password=''
#전송 정보
smtp_info = dict({"smtp_server" : "smtp.naver.com", # SMTP 서버 주소
                  "smtp_user_id" : id,
                  "smtp_user_pw" : password ,
                  "smtp_port" : 587}) # SMTP 서버 포트

# 메일 내용 작성
title = "Netflex Top 10 영화 & TV 워드클라우드"
content = "Netflex Top 10 영화 & TV 워드클라우드 전송드립니다."
sender = "hahxowns@naver.com"
receiver = "hahxowns@gmail.com"


msg = MIMEText(_text = content, _charset = "utf-8") # 메일 내용

# 첨부파일 추가
multi = make_multimsg(attach_dict)
multi['subject'] = title    # 메일 제목
multi['from'] = sender      # 송신자
multi['to'] = receiver      # 수신자
multi.attach(msg)

send_email(smtp_info, multi)