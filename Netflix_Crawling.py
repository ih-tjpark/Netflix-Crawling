import smtplib
from email.mime.multipart import MIMEMultipart
from email import encoders
import pandas as pd

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
            with open(value['filename'], 'rb') as fp:
                msg = MIMEImage(fp.read(), _subtype=value['subtype'])
        elif key == 'audio':
            with open(value['filename'], 'rb') as fp:
                msg = MIMEAudio(fp.read(), _subtype=value['subtype'])
        else:
            with open(value['filename'], 'rb') as fp:
                msg = MIMEBase(value['maintype'],  _subtype=value['subtype'])
                msg.set_payload(fp.read())
                encoders.encode_base64(msg)
        # 파일 이름을 첨부파일 제목으로 추가
        msg.add_header('Content-Disposition', 'attachment', filename=value['filename'])

        # 첨부파일 추가
        multi.attach(msg)

    return multi

#전송 정보
smtp_info = dict({"smtp_server" : "smtp.naver.com", # SMTP 서버 주소
                  "smtp_user_id" : "hahxowns@naver.com",
                  "smtp_user_pw" : "wnsxo123" ,
                  "smtp_port" : 587}) # SMTP 서버 포트

# 메일 내용 작성
title = "기본 이메일 입니다."
content = "메일 내용입니다."
sender = "hahxowns@naver.com"
receiver = "hahxowns@gmail.com"

'''
attach_dict = {
    'text' : {'maintype' : 'text', 'subtype' :'plain', 'filename' : 'test.txt'}, # 텍스트 첨부파일
    'image' : {'maintype' : 'image', 'subtype' :'jpg', 'filename' : 'image1.png' }, # 이미지 첨부파일
    'audio' : {'maintype' : 'audio', 'subtype' :'mp3', 'filename' : 'test.mp3' }, # 오디오 첨부파일
    'video' : {'maintype' : 'video', 'subtype' :'mp4', 'filename' : 'test.mp4' }, # 비디오 첨부파일
    'application' : {'maintype' : 'application', 'subtype' : 'octect-stream', 'filename' : 'test.pdf'} # 그외 첨부파일
}
'''
attach_dict = {
    'image' : {'maintype' : 'image', 'subtype' :'jpg', 'filename' : 'image1.png' }, # 이미지 첨부파일
    'application' : {'maintype' : 'application', 'subtype' : 'octect-stream', 'filename' : 'test.pdf'}
}
excel = pd.DataFrame({

})

msg = MIMEText(_text = content, _charset = "utf-8") # 메일 내용

# 첨부파일 추가
multi = make_multimsg(attach_dict)
multi['subject'] = title    # 메일 제목
multi['from'] = sender      # 송신자
multi['to'] = receiver      # 수신자
multi.attach(msg)

send_email(smtp_info, multi)
