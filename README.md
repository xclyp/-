### **第一阶段：环境与文件准备**

在开始写代码之前，我们需要准备好电脑环境和所需的文件。

#### **步骤 1：安装 Python**

如果您的电脑还没有安装Python，需要先安装它。

1.  访问 Python 官网：[https://www.python.org/downloads/](https://www.python.org/downloads/)
2.  下载最新的稳定版本（例如 Python 3.11 或更高版本）。
3.  运行安装程序。**特别重要的一步**：在安装界面的最下方，**务必勾选 "Add Python to PATH"** 这个选项，这样可以简化后续操作。然后点击 "Install Now" 即可。

#### **步骤 2：安装 `pandas` 库**

我们需要一个库来方便地读取我们创建的编辑信息表。

1.  打开您电脑的命令行工具：
    *   **Windows**: 按 `Win + R` 键，输入 `cmd`，然后按回车。
    *   **Mac**: 打开 "终端" (Terminal) 应用。
2.  在打开的黑色窗口中，输入以下命令，然后按回车：
    ```bash
    pip install pandas
    ```
    等待它自动下载并安装完成。

#### **步骤 3：创建一个项目文件夹**

为了保持所有文件整洁，在您的电脑上（比如桌面）创建一个新的文件夹，命名为 `AutoSubmission` 或任何您喜欢的名字。**之后的所有文件都将放在这个文件夹里**。

#### **步骤 4：创建 `editors.csv` 编辑信息表**

这是您的投递名单。

1.  打开 Excel 或 Google Sheets。
2.  按照下面的格式创建表头和内容。**表头名称必须和示例完全一样**，因为代码会根据这些名称来读取数据。
    | 编辑姓名 | 邮箱地址 | 媒体/期刊名称 | 尊称 |
    | :--- | :--- | :--- | :--- |
    | 张三 | zhangsan@example.com | 科技新观察 | 张编辑 |
    | 李四 | lisi@example.com | 文学月刊 | 李老师 |
    | 王五 | wangwu@example.com | 商业周报 | 王主编 |
    | ... | ... | ... | ... |
3.  将这个文件另存为 **CSV (逗号分隔)** 格式。在保存时，文件名命名为 `editors.csv`，并把它**保存到您刚刚创建的 `AutoSubmission` 文件夹中**。
    *   **注意**：如果您的表格中有中文，保存时请确保编码为 `UTF-8`，以防乱码。

#### **步骤 5：创建 `template.txt` 邮件模板**

这是您的邮件正文模板。

1.  在 `AutoSubmission` 文件夹中，右键新建一个文本文档。
2.  将它命名为 `template.txt`。
3.  打开这个文件，粘贴并修改以下模板内容。其中的 `{}` 是占位符，程序会自动替换它们。
    ```text
    尊敬的{尊称}：

    您好！

    我是[您的名字]，长期关注贵刊《{媒体/期刊名称}》。我的稿件《您的文章标题》与贵刊的[某个栏目]定位非常契合，特此投稿，希望能得到您的审阅。

    稿件已添加至附件，请您查收。

    期待您的回复！

    祝好！

    [您的名字]
    [您的联系方式]
    [日期]
    ```
4.  保存并关闭文件。

#### **步骤 6：准备稿件附件**

将您要投递的稿件文件（例如 `我的稿件.docx` 或 `我的稿件.pdf`）也**复制一份到 `AutoSubmission` 文件夹中**。

---

### **第二阶段：邮箱设置**

这一步至关重要，是为了让您的代码能够安全地登录您的邮箱并发送邮件。

#### **步骤 7：开启 SMTP 服务并获取授权码**

**警告：** 代码中使用的不是您的邮箱登录密码，而是一个专用的“授权码”！

您需要登录您的邮箱网页版，进行以下设置（以Gmail和QQ邮箱为例）：

*   **对于 Gmail:**
    1.  您必须先开启“两步验证”(2-Step Verification)。
    2.  然后访问 [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)。
    3.  在“选择应用”中选“邮件”，在“选择设备”中选“Windows计算机”或“其他”，然后点击“生成”。
    4.  它会生成一串16位的密码，**这就是您的授权码**。请立即复制并保存好它，因为这个窗口关闭后就看不到了。

*   **对于 QQ 邮箱:**
    1.  登录后，进入“设置” -> “账户”。
    2.  向下滚动找到 "POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务"。
    3.  确保 "IMAP/SMTP服务" 是开启状态，然后点击“生成授权码”。
    4.  根据提示发送短信，然后您会得到一串字符，**这就是您的授权码**。

*   **对于 163 邮箱:**
    1.  登录后，进入“设置” -> “客户端授权密码”。
    2.  开启服务，并设置一个授权码。

---

### **第三阶段：编写并运行代码**

现在，所有准备工作都已完成，我们可以开始创建并运行脚本了。

#### **步骤 8：创建 `send_emails.py` Python脚本**

1.  在 `AutoSubmission` 文件夹中，再次新建一个文本文档。
2.  将它重命名为 `send_emails.py` (注意后缀是 `.py`)。
3.  用记事本或任何代码编辑器打开它，将下面的全部代码复制粘贴进去。

```python
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
EMAIL_ADDRESS = 'your_email@qq.com'  # 你的邮箱地址
EMAIL_PASSWORD = 'your_app_password_here'  # 这里粘贴你获取的16位授权码，不是登录密码！

# -- 文件路径配置 --
CSV_FILE_PATH = 'editors.csv'  # 编辑信息CSV文件
TEMPLATE_FILE_PATH = 'template.txt' # 邮件正文模板
ATTACHMENT_FILENAME = '我的稿件.docx'  # 你的稿件文件名，确保它和文件真实名字一致

# --- 2. 读取邮件模板 ---
try:
    with open(TEMPLATE_FILE_PATH, 'r', encoding='utf-8') as f:
        email_template = f.read()
except FileNotFoundError:
    print(f"错误：找不到邮件模板文件 {TEMPLATE_FILE_PATH}。请检查文件名和路径。")
    exit()

# --- 3. 读取编辑列表 ---
try:
    df = pd.read_csv(CSV_FILE_PATH, encoding='utf-8')
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
                媒体/期刊名称=publication
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

```

#### **步骤 9：修改代码中的配置**

打开 `send_emails.py` 文件，找到顶部的 `# --- 1. 请在这里修改你的配置信息 ---` 部分，将 `SMTP_SERVER`, `SMTP_PORT`, `EMAIL_ADDRESS`, `EMAIL_PASSWORD` 和 `ATTACHMENT_FILENAME` 的值**替换成你自己的信息**。

#### **步骤 10：运行脚本！**

1.  再次打开命令行工具 (cmd 或 Terminal)。
2.  使用 `cd` 命令进入到你的项目文件夹。例如，如果你的文件夹在桌面上，可以输入：
    ```bash
    cd Desktop/AutoSubmission
    ```
3.  一切就绪后，输入以下命令并按回车，程序就会开始运行：
    ```bash
    python send_emails.py
    ```

如果一切顺利，您会看到命令行窗口中逐条打印出“成功发送邮件给 XXX”的信息。同时，您可以登录邮箱的“已发送”文件夹查看，邮件已经按照您的模板，带着附件，一封封地发送出去了！

### **常见问题排查**

*   **登录失败 (AuthenticationError)**：99% 的可能是你的 `EMAIL_PASSWORD` 填的是登录密码，而不是**授权码**。请检查步骤7。
*   **找不到文件 (FileNotFoundError)**：请确保 `editors.csv`, `template.txt`, `我的稿件.docx` 这三个文件都和 `send_emails.py` 在**同一个文件夹**下，并且文件名与代码中的配置完全一致。
*   **连接超时 (Timeout)**：可能是 `SMTP_SERVER` 或 `SMTP_PORT` 填错了，或者你的网络防火墙阻止了连接。可以尝试更换端口（比如587换成465）。
