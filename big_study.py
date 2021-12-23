from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import xlrd
import numpy
from multiprocessing.dummy import Pool
from selenium.webdriver.chrome.options import Options
import os
from smtplib import SMTP_SSL
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium.webdriver.support.select import Select



def sendMail(message, Subject, sender_show, recipient_show, to_addrs, cc_show=''):
    """
    :param message: str 邮件内容
    :param Subject: str 邮件主题描述
    :param sender_show: str 发件人显示，不起实际作用如："xxx"
    :param recipient_show: str 收件人显示，不起实际作用 多个收件人用','隔开如："xxx,xxxx"
    :param to_addrs: str 实际收件人
    :param cc_show: str 抄送人显示，不起实际作用，多个抄送人用','隔开如："xxx,xxxx"
    """
    # 填写真实的发邮件服务器用户名、授权码
    user = '******@qq.com'
    password = '*******8'
    # 邮件内容
    msg = MIMEMultipart()
    msg.attach(MIMEText(message, 'html', _charset="utf-8"))

    path = 'img/'
    imgs = os.listdir(path=path)
    print(imgs)
    i=0
    for img in imgs:
        # 构造附件1
        print(img)
        i=i+1
        att = MIMEText(open('img/'+img, 'rb').read(), 'base64', 'utf-8')
        att["Content-Type"] = 'application/octet-stream'
        # 附件名称为中文时的写法
        att.add_header("Content-Disposition", "attachment", filename=("gbk", "", img))
        # 附件名称非中文时的写法,这里的filename可以任意写，写什么名字，邮件中显示什么名字
        # att["Content-Disposition"] = 'attachment; filename="{}"'.format(filename)
        msg.attach(att)
    # 邮件主题描述
    msg["Subject"] = Subject
    # 发件人显示，不起实际作用
    msg["from"] = sender_show
    # 收件人显示，不起实际作用
    msg["to"] = recipient_show
    # 抄送人显示，不起实际作用
    msg["Cc"] = cc_show
    print('准备发送截图')
    with SMTP_SSL(host="smtp.qq.com", port=465) as smtp:
        # 登录发送邮件服务器
        smtp.login(user=user, password=password)
        # print('登陆成功')
        # 实际发送、接收邮件配置
        smtp.sendmail(from_addr=user, to_addrs=to_addrs.split(','), msg=msg.as_string())
        print('发送成功')

def one_study(index):
    tag = True
    i=index
    chrome_options = Options()
    chrome_options.add_argument("--window-size=1920,1080")
    chrome = webdriver.Chrome(options=chrome_options)
    chrome.get('https://jinshuju.net/f/la7dt8')
    #print('正在打开首页')
    while tag:
        try:
            input_list = chrome.find_elements(by=By.CLASS_NAME,value='ant-input')
            tag=False
        except:
            time.sleep(1)
    tag=True

    #print('正在填写姓名')
    while tag:
        try:
            input_name=chrome.find_element(by=By.NAME,value='field_1')
            input_name.send_keys(sheet1.cell(i,0).value)
            tag=False
        except:
            time.sleep(1)
    tag=True
    #print('正在填写学号')
    while tag:
        try:
            input_id=chrome.find_element(by=By.NAME,value='field_2')
            input_id.send_keys(int(sheet1.cell(i,1).value))
            tag=False
        except:
            time.sleep(1)
    tag=True
    #print('正在填写手机号')
    while tag:
        try:
            input_tel=chrome.find_element(by=By.NAME,value='field_3')
            input_tel.send_keys(int(sheet1.cell(i,2).value))
            tag=False
        except:
            time.sleep(1)
    tag=True
    #print('正在填写专业')
    while tag:
        try:
            input_major=chrome.find_element(by=By.NAME,value='field_34')
            input_major.send_keys(sheet1.cell(i,3).value)
            tag=False
        except:
            time.sleep(1)
    tag=True
    #print('正在选择学院')
    while tag:
        try:
            chrome.find_element(by=By.CLASS_NAME,value='pretty-select__value-container').click()
            tag=False
        except:
            time.sleep(1)
    tag=True

    while tag:
        try:
            # chrome.find_element(by=By.ID,value='react-select-2-option-21').click()
            schools = chrome.find_elements(by=By.CLASS_NAME, value='pretty-select__option')
            for school in schools:
                if str(sheet1.cell(i,4).value) in school.text:
                    school.click()
                    tag = False
                    break
        except:
            time.sleep(1)
    tag=True
    #print('正在选择团支部')
    while tag:
        try:
            chrome.find_elements(by=By.CLASS_NAME,value='pretty-select__value-container')[1].click()
            tag=False
        except:
            time.sleep(1)
    tag=True

    while tag:
        try:
            sub_minist = chrome.find_elements(by=By.CLASS_NAME, value='pretty-select__option')
            for sub in sub_minist:
                if str(sheet1.cell(i, 5).value) in sub.text:
                    #print(sub.text)
                    sub.click()
                    tag = False
                    break
        except:
            time.sleep(1)
    tag=True
    #print('正在点击提交')
    while tag:
        try:
            chrome.find_element(by=By.CLASS_NAME,value='ant-btn').click()
            tag=False
        except:
            time.sleep(1)
    tag=True

    #print('正在登录青年大学习账户')
    while tag:
        try:
            chrome.find_element(by=By.CLASS_NAME,value='yonghu1').click()
            tag=False
        except:
            time.sleep(1)
    tag=True

    #print('正在输入青年大学习账号（手机号）')
    while tag:
        try:
            study_username=chrome.find_element(by=By.ID,value='login_username')
            study_username.send_keys(int(sheet1.cell(i,6).value))
            tag=False
        except:
            time.sleep(1)
    tag=True

    #print('正在输入青年大学习密码')
    while tag:
        try:
            study_password=chrome.find_element(by=By.ID,value='login_password')
            password=str(sheet1.cell(i,7).value)
            #print(password)
            study_password.send_keys(password)
            tag=False
        except:
            time.sleep(1)
    tag=True

    #print('正在点击登录')
    while tag:
        try:
            chrome.find_element(by=By.ID,value='login_btn').click()
            tag=False
        except:
            time.sleep(1)
    tag=True

    while tag:
        try:
            chrome.find_element(by=By.CLASS_NAME,value='layui-layer-btn0').click()
            tag=False
        except:
            time.sleep(1)
    tag=True

    while tag:
        try:
            chrome.find_element(by=By.CLASS_NAME,value='yonghu1').click()
            tag=False
        except:
            time.sleep(1)
    tag=True
    #print('正在切换iframe')
    while tag:
        try:
            frames = chrome.find_elements(by=By.TAG_NAME,value='iframe')
            chrome.switch_to.frame(frames[0])
            tag=False
        except:
            time.sleep(1)
    tag=True
    #print('正在点击province')
    while tag:
        try:
            province_select=chrome.find_element(by=By.ID,value='province')
            province_select.click()
            tag=False
        except:
            time.sleep(1)
    tag=True

    while tag:
        try:
            chrome.find_element(by=By.XPATH,value='/html/body/div[2]/select[1]/option[26]').click()
            chrome.find_element(by=By.XPATH,value='/html/body/div[2]/select[1]/option[26]').click()
            tag = False
        except:
            time.sleep(1)
    tag=True
    #print('正在点击city')
    while tag:
        try:
            chrome.find_element(by=By.ID,value='city').click()
            tag=False
        except:
            time.sleep(1)
    tag=True

    while tag:
        try:
            chrome.find_element(by=By.XPATH,value='/html/body/div[2]/select[2]/option[2]').click()
            chrome.find_element(by=By.ID,value='city').click()
            tag = False
        except:
            time.sleep(1)
    tag = True
    while tag:
        try:
            chrome.find_element(by=By.CLASS_NAME,value='sure').click()
            tag = False
        except:
            time.sleep(1)
    tag = True
    #print('正在开始学习')
    while tag:
        try:
            chrome.find_element(by=By.CLASS_NAME,value='start_btn').click()
            tag=False
        except:
            time.sleep(1)
    tag=True
    strat_time = time.time()
    #print('正在试图回答问题')
    while tag:
        try:
            try:
                continues = chrome.find_elements(by=By.CLASS_NAME, value='continue')
                for continuei in continues:
                    try:
                        continuei.click()
                    except:
                        pass
            except:
                pass
            w0_list = chrome.find_elements(by=By.CLASS_NAME,value='w0')
            button_list = chrome.find_elements(by=By.CLASS_NAME,value='button')
            for j in range(0,len(w0_list)):
                w0=w0_list[j]
                try:
                    w0.click()
                except:
                    pass
            for j in range(0,len(button_list)):
                button=button_list[j]
                try:
                    button.click()
                except:
                    pass

            if time.time()-strat_time>=600:
                tag=False
        except:
            time.sleep(1)
    tag = True
    #print('播放完成')

    chrome.back()
    while tag:
        try:
            chrome.find_element(by=By.CLASS_NAME, value='nv3 ').click()
            tag=False
        except:
            time.sleep(1)
    tag=True
    time.sleep(7)
    img_name=sheet1.cell(i,0).value
    file_name='.\\img\\'+img_name+'.png'
    # chrome.maximize_window()
    chrome.get_screenshot_as_file(file_name)

    font = ImageFont.truetype(".\\荆南麦圆体.ttf", 56)
    imageFile = f".\\img\\{img_name}.png"
    im1 = Image.open(imageFile)

    #画图
    draw = ImageDraw.Draw(im1)
    draw.text((800, 500),img_name, (255, 0, 0), font=font)    #设置文字位置/内容/颜色/字体
    draw = ImageDraw.Draw(im1)                          #Just draw it!

    #另存图片
    im1.save(f".\\img\\{img_name}.png")
    chrome.quit()



xls = xlrd.open_workbook_xls('information.xls')
names = xls.sheet_names()
sheet1 = xls.sheet_by_name(names[0])
nrows = sheet1.nrows
try:
    os.makedirs('img')
except:
    pass
sb_list = numpy.linspace(1,nrows-1,nrows-1,dtype=int)
pool = Pool(5)
pool.map(one_study,sb_list)
pool.close()
pool.join

message = '''
<p>你好！团支书，这是这周的大学习截图</p>
<p>请查收</p>
'''
Subject = '青年大学习截图'
# 显示发送人
sender_show = '*******@qq.com'
# 显示收件人
recipient_show = '*******@qq.com'
# 实际发给的收件人
to_addrs = '*********@qq.com'
sendMail(message, Subject, sender_show, recipient_show, to_addrs)


