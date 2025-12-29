import imaplib
import email
import datetime
import re
import os
from email.header import decode_header

# 连接到QQ邮箱IMAP服务器
mail = imaplib.IMAP4_SSL("imap.qq.com")
mail.login("您的QQ号@qq.com", "您的授权码")

mail.select("inbox")

# 1. 设定搜索日期范围
start_date = datetime.date(2025, 12, 28) # 起始日期
end_date = None # 结束日期 (None 表示直到现在)
# end_date = datetime.date(2025, 12, 25) # 如果要指定结束日期，取消注释并修改


months = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}

def get_imap_date(d):
    return f'{d.day}-{months[d.month]}-{d.year}'

start_str = get_imap_date(start_date)
search_query = f'SINCE "{start_str}"'

if end_date:
    # IMAP BEFORE 是不包含该日期的，所以为了包含 end_date，需要加一天
    cutoff_date = end_date + datetime.timedelta(days=1)
    end_str = get_imap_date(cutoff_date)
    search_query += f' BEFORE "{end_str}"'
    print(f"正在搜索 {start_str} 到 {get_imap_date(end_date)} (包含) 的所有邮件...")
else:
    print(f"正在搜索 {start_str} 以来(包含该日)的所有邮件...")

# 2. 搜索指定日期范围的邮件
status, messages = mail.search(None, search_query)
mail_ids = messages[0].split()

# 倒序，从最新开始处理
target_ids = list(reversed(mail_ids))

print(f"共找到 {len(target_ids)} 封邮件。\n")

# 确保下载目录存在
download_dir = "downloads"
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

# 3. 循环处理每一封邮件
for i, email_id in enumerate(target_ids):
    print(f"正在读取第 {i+1} 封邮件 (ID: {email_id.decode()})...")
    
    # 为了下载附件，我们需要获取邮件内容
    # 虽然 (RFC822) 会下载整封邮件，但这是获取附件内容最通用的方式
    status, msg_data = mail.fetch(email_id, "(RFC822)")

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            
            # --- 解析主题 ---
            subject_header = msg["Subject"]
            if subject_header:
                decoded_list = decode_header(subject_header)
                subject_bytes, encoding = decoded_list[0]
                if isinstance(subject_bytes, bytes):
                    # 如果没有指定编码，默认使用 utf-8
                    subject = subject_bytes.decode(encoding if encoding else "utf-8", errors="replace")
                else:
                    subject = subject_bytes # 已经是字符串
            else:
                subject = "无主题"
            
            print(f"  主题: {subject}")

            # --- 确定分类文件夹 ---
            # 辅助函数：尝试从文本中提取分类
            def get_category(text):
                if not text:
                    return None
                
                # 中文数字映射
                cn_map = {'一': '1', '二': '2', '三': '3', '四': '4', '五': '5', 
                          '六': '6', '七': '7', '八': '8', '九': '9', '十': '10'}
                
                def to_arabic(num_str):
                    return cn_map.get(num_str, num_str)

                # 辅助：判断是否为合理的作业编号 (避免匹配到学号)
                def is_valid_num(num_str):
                    # 如果是数字字符串，长度不应超过2位 (防止匹配到学号/年份)
                    if num_str.isdigit():
                        return len(num_str) <= 2
                    return True

                # 检查实验报告 (LAB)
                # 1. 匹配 "第X次...实验" 格式 (允许中间有其他字符)
                lab_match_1 = re.search(r'第\s*(\d+|[一二三四五六七八九十]+)\s*次.*?(?:实验|LAB|实践)', text, re.IGNORECASE)
                if lab_match_1:
                    return f"LAB{to_arabic(lab_match_1.group(1))}"
                
                # 2. 匹配 "实验报告1", "LAB 1", "实验作业2", "实验4" 等 (数字在后)
                lab_match_2 = re.search(r'(?:实验报告|实验作业|实践|LAB|实验).*?(\d+|[一二三四五六七八九十]+)', text, re.IGNORECASE)
                if lab_match_2 and is_valid_num(to_arabic(lab_match_2.group(1))):
                    return f"LAB{to_arabic(lab_match_2.group(1))}"
                
                # 检查课堂作业
                # 1. 匹配 "第X次...作业" 格式 (允许中间有其他字符)
                class_match_1 = re.search(r'第\s*(\d+|[一二三四五六七八九十]+)\s*次.*?(?:课堂作业|作业)', text, re.IGNORECASE)
                if class_match_1:
                    return f"课堂作业{to_arabic(class_match_1.group(1))}"

                # 2. 匹配 "课堂作业1", "作业 2" 等 (数字在后)
                class_match_2 = re.search(r'(?:课堂作业|作业).*?(\d+|[一二三四五六七八九十]+)', text, re.IGNORECASE)
                if class_match_2 and is_valid_num(to_arabic(class_match_2.group(1))):
                    return f"课堂作业{to_arabic(class_match_2.group(1))}"
                
                return None

            folder_from_subject = get_category(subject)

            # --- 遍历查找并下载附件 ---
            attachment_count = 0
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                
                fileName = part.get_filename()
                if bool(fileName):
                    attachment_count += 1
                    # 解析文件名
                    decoded_list = decode_header(fileName)
                    fname_bytes, encoding = decoded_list[0]
                    
                    if isinstance(fname_bytes, bytes):
                        try:
                            fileName = fname_bytes.decode(encoding if encoding else "utf-8")
                        except (UnicodeDecodeError, LookupError):
                            try:
                                fileName = fname_bytes.decode("gbk")
                            except UnicodeDecodeError:
                                fileName = fname_bytes.decode("gb18030", errors="replace")
                    
                    # 处理文件名中的非法字符 (防止路径错误)
                    fileName = "".join([c for c in fileName if c not in r'\/:*?"<>|'])

                    # --- 优先基于原始文件名提取分类 (避免重命名引入主题中的混淆信息) ---
                    folder_from_file = get_category(fileName)

                    # --- 尝试修复无意义的文件名 ---
                    # 如果文件名看起来像哈希值(例如QQ截图生成的)或通用名称，且主题包含有效信息，则尝试用主题重命名
                    name_root, name_ext = os.path.splitext(fileName)
                    # 匹配 16位以上纯十六进制/数字字母组合，或 image/screenshot/wx_camera/新建 开头
                    is_garbage_name = re.match(r'^([a-fA-F0-9]{16,}|image|screenshot|wx_camera|新建|IMG|IMAGE)', name_root, re.IGNORECASE)
                    
                    if is_garbage_name and subject and subject != "无主题":
                        safe_subject = "".join([c for c in subject if c not in r'\/:*?"<>|'])
                        if safe_subject:
                            new_fileName = f"{safe_subject}{name_ext}"
                            print(f"    [重命名] 检测到无意义文件名，已重命名: {fileName} -> {new_fileName}")
                            fileName = new_fileName
                    
                    # --- 补充信息：如果文件名缺失关键信息（如学号），尝试用主题补全 ---
                    # 提取主题中的学号 (8位以上数字)
                    subj_id_match = re.search(r'\d{8,}', subject)
                    subj_id = subj_id_match.group(0) if subj_id_match else None
                    
                    should_rename = False
                    # 情况1: 主题包含学号，但文件名不包含
                    if subj_id and subj_id not in fileName:
                        should_rename = True
                    # 情况2: 文件名是通用名称 (如 "第二次作业.docx") 且不包含长数字
                    elif not re.search(r'\d{8,}', fileName):
                        if re.search(r'(?:作业|实验|报告|文档|DOCX|PDF)', fileName, re.IGNORECASE):
                             should_rename = True

                    if should_rename and subject and subject != "无主题":
                        safe_subject = "".join([c for c in subject if c not in r'\/:*?"<>|'])
                        # 避免重复：如果文件名已经是主题的一部分，或者主题包含文件名，需谨慎
                        if safe_subject not in fileName: 
                             new_fileName = f"{safe_subject}_{fileName}"
                             print(f"    [重命名] 补充身份信息: {fileName} -> {new_fileName}")
                             fileName = new_fileName

                    # --- 检查文件扩展名 ---
                    # 只允许 pdf, doc, docx, jpg, png
                    allowed_extensions = ('.pdf', '.doc', '.docx', '.jpg', '.png')
                    if not fileName.lower().endswith(allowed_extensions):
                        print(f"    [跳过] 不支持的文件格式: {fileName}")
                        continue

                    # --- 智能分类逻辑 ---
                    # folder_from_file 已经在重命名之前计算过了
                    
                    target_folder = "tmp" # 默认存入 tmp (有疑问/未分类)
                    
                    if folder_from_subject and folder_from_file:
                        if folder_from_subject == folder_from_file:
                            target_folder = folder_from_subject
                        else:
                            # 冲突：主题和文件名分类不一致 -> 优先根据文件名进行分类
                            print(f"    [警告] 分类冲突: 主题({folder_from_subject}) vs 文件名({folder_from_file}) -> 优先使用文件名")
                            target_folder = folder_from_file
                    elif folder_from_subject:
                        target_folder = folder_from_subject
                    elif folder_from_file:
                        target_folder = folder_from_file
                    else:
                        # 均未匹配到分类 -> 存入 tmp
                        target_folder = "tmp"
                    
                    # 创建分类目录
                    target_dir = os.path.join(download_dir, target_folder)
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir)
                    
                    # 下载附件
                    filepath = os.path.join(target_dir, fileName)
                    try:
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        print(f"    [附件] 已下载至 [{target_folder}]: {fileName}")
                    except Exception as e:
                        print(f"    [附件] 下载失败: {fileName} ({e})")

            if attachment_count == 0:
                print("    [无附件]")
            
            print("-" * 30)

mail.close()
mail.logout()