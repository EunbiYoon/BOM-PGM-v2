import os
import pandas as pd
import numpy as np
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 

title='[테네시 재료비 관리 Task] 6월 4주차 BOM과 실제 생산 투입 재료비 차이 분석'
################################## Send email ################################## 
server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
msg['To']='iggeun.kwon@lge.com, incheol.kang@lge.com, aaron1.garcia@lge.com, jacey.jung@lge.com, gilnam.lee@lge.com, steven.yang@lge.com, jajoon1.koo@lge.com, wolyong.ha@lge.com, dowan.han@lge.com, grace.hwang@lge.com'
msg['Cc']='ethan.son@lge.com, jongseop.kim@lge.com, richard.song@lge.com, minhyoung.sun@lge.com, kitae3.park@lge.com, tg.kim@lge.com'
msg['Bcc']='eunbi1.yoon@lge.com'

#title
msg['Subject']=title
msg.attach(MIMEText('<h4 style="font-family:Arial Narrow; font-weight:500">Dear All,<br/><br/>I would like to share this week bom comparison report and detailed informations are in below website.<br/>You can access website with Chrome or Edge in CloudPC or LG wifi for security purpose. <a href="http://10.225.2.85">http://10.225.2.85</a></h4>','html'))

save_path='C:/Users/RnD Workstation/Documents/NPTGERP/0623/'

# graph file read
with open(save_path+'FL_0623.png', 'rb') as f:
        img_data = f.read()
image2 = MIMEImage(img_data, name=os.path.basename("FL.png"))

with open(save_path+'TL_0623.png', 'rb') as f:
        img_data = f.read()
image1 = MIMEImage(img_data, name=os.path.basename("TL.png"))

with open(save_path+'DR_0623.png', 'rb') as f:
        img_data = f.read()
image3 = MIMEImage(img_data, name=os.path.basename("DR.png"))



# msg.attach(MIMEText('<br/><h3 style="font-family:sans-serif;">Front Loader BPA Entity Trend</h3>','html'))
msg.attach(image1)
msg.attach(image2)
msg.attach(image3)


#첨부 파일
etcFileName='Cost Review_0623.xlsx'
with open("C:/Users/RnD Workstation/Documents/NPTGERP/0623/result_0623.xlsx", 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)

etcFileName='BOM Comparison_FL_0623.xlsx'
with open("C:/Users/RnD Workstation/Documents/NPTGERP/0623/BOM Comparison_FL.xlsx", 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)

etcFileName='BOM Comparison_TL_0623.xlsx'
with open("C:/Users/RnD Workstation/Documents/NPTGERP/0623/BOM Comparison_TL.xlsx", 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)

etcFileName='BOM Comparison_DR_0623.xlsx'
with open("C:/Users/RnD Workstation/Documents/NPTGERP/0623/BOM Comparison_DR.xlsx", 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)

#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")
