import smtplib
import pandas as pd
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# --- 1. 请在这里修改你的配置信息 ---
# -- 邮箱服务器配置 --
# (常见邮箱SMTP服务器: QQ: smtp.qq.com, 163: smtp.163.com, Gmail: smtp.gmail.com)
SMTP_SERVER = 'smtp.qq.com'
# (常见端口: SSL端口465, TLS端口587。如果一个不行就换另一个)
SMTP_PORT = 465
# -- 你的邮箱账户信息 --
EMAIL_ADDRESS = '2563688491@qq.com'  # 你的邮箱地址
EMAIL_PASSWORD = 'trmppiptwkpzdigj'  # 这里粘贴你获取的16位授权码，不是登录密码！

# -- 文件路径配置 --
CSV_FILE_PATH = 'editors.csv'  # 编辑信息CSV文件
TEMPLATE_FILE_PATH = 'template.txt' # 邮件正文模板
ATTACHMENT_FILENAME = '知乎短篇.txt'  # 你的稿件文件名，确保它和文件真实名字一致

# --- 2. 读取邮件模板 ---
try:
    with open(TEMPLATE_FILE_PATH, 'r', encoding='utf-8') as f:
        email_template = f.read()
except FileNotFoundError:
    print(f"错误：找不到邮件模板文件 {TEMPLATE_FILE_PATH}。请检查文件名和路径。")
    exit()

# --- 3. 读取编辑列表 ---
# --- 3. 读取编辑列表 ---
try:
    df = pd.read_csv(CSV_FILE_PATH, encoding='utf-8')
# <--- 把 utf-8 改成 gbk
except FileNotFoundError:

    print(f"错误：找不到编辑列表文件 {CSV_FILE_PATH}。请检查文件名和路径。")
    exit()

# --- 4. 检查附件是否存在 ---
if not os.path.exists(ATTACHMENT_FILENAME):
    print(f"错误：找不到附件 {ATTACHMENT_FILENAME}。请检查文件名和路径。")
    exit()

# --- 5. 连接到SMTP服务器并发送邮件 ---
try:
    # 根据端口选择不同的连接方式
    if SMTP_PORT == 465:
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    else: # 587 or other
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()

    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    print("成功登录邮箱...")

    # --- 遍历列表并发送邮件 ---
    for index, row in df.iterrows():
        try:
            # 从表格中获取信息
            editor_name = row['编辑姓名']
            editor_email = row['邮箱地址']
            publication = row['媒体/期刊名称']
            salutation = row['尊称']

            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = f"你的名字 <{EMAIL_ADDRESS}>" # 显示发件人名称
            msg['To'] = editor_email
            msg['Subject'] = f"投稿：关于《您的文章标题》- 来自 [您的名字]"

            # 个性化邮件正文
            body = email_template.format(
                尊称=salutation,
                publication_name=publication
            )
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # 添加附件
            with open(ATTACHMENT_FILENAME, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            # 处理附件文件名中的中文
            part.add_header('Content-Disposition', 'attachment', filename=('gbk', '', ATTACHMENT_FILENAME))
            msg.attach(part)

            # 发送邮件
            server.send_message(msg)
            print(f"[{index + 1}/{len(df)}] 成功发送邮件给 {editor_name} ({editor_email})")

            # 友好等待，避免被服务器认为是垃圾邮件攻击
            time.sleep(5) # 每发送一封后等待5秒

        except Exception as e:
            print(f"发送给 {row.get('编辑姓名', '未知收件人')} 时发生错误: {e}")


except smtplib.SMTPAuthenticationError:
    print("邮箱登录失败！请检查：1. 邮箱地址是否正确。2. 是否使用了授权码而非登录密码。3. 授权码是否已过期。")
except Exception as e:
    print(f"发生未知错误: {e}")
finally:
    if 'server' in locals() and server:
        server.quit()
        print("已断开与邮件服务器的连接。")

