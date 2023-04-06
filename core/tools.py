import csv
import openpyxl
from openpyxl.styles import Font
import os
import re
from core import my_errors
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText


def read_csv(csv_name, encoding='utf-8-sig'):
    if not os.path.exists(csv_name):
        raise my_errors.NoFile(csv_name)
    csv_file = open(csv_name, 'r', encoding=encoding)
    return csv.reader(csv_file)


def write_csv(csv_name, encoding='utf-8-sig', newline=''):
    csv_file = open(csv_name, 'w', encoding=encoding, newline=newline)
    return csv.writer(csv_file)


# 用于判断学生是否完成青年大学习
def finished(student: list, table) -> bool:  # student是学生信息，table是已完成名单
    # student的0，1，2分别代表姓名、学号、支部
    ret = False
    length = len(table)
    index = 0
    while index < length:
        # 当某个班级中某个学生的姓名或者学号在已完成名单里，执行下面if语句
        # 考虑到会出现重名的情况（最极端的情况是同一个班级里有数个名字一样的人），同时会对学号和班级信息进行判断
        # judge = (table[index][4] == student[1] or table[index][4] == student[0]) \
        #         and table[index][6] == student[2]
        table_str = ' '.join(table[index])
        judge = re.findall(student[0], table_str) and re.findall(student[2], table_str) or \
                re.findall(student[1], table_str) and re.findall(student[2], table_str)
        # judge = student[0] in table[index] and student[2] in table[index] or \
        #          student[1] in table[index] and student[2] in table[index]
        if judge:
            ret = True
            break
        index += 1
    return ret


# 把列表转化为excel的工作表，返回一个工作表
def list_to_sheet(my_list: list, sheet, font=Font(name='等线')):
    row = 1
    for i in my_list:
        column = 1
        for j in i:
            sheet.cell(row, column).value = j
            sheet.cell(row, column).font = font
            column += 1
        row += 1
    return sheet


# 用于存储excel文件，如果不声明存储目录默认为文件根目录
def save_excel(file: openpyxl.Workbook, file_name, save_dir='.\\'):
    if not os.path.exists(save_dir):
        os.mkdir(save_dir)
    file.save(os.path.join(save_dir, file_name))


# 按照文件导出时间进行归类
def classify_file(file, sort_file):
    file_path = ''
    return file_path


# TODO 记录原始表格导出日期
# FIXME 使用文件修改时间作为导出日期，在恶意修改文件以及其他情况下，可能会导致bug
def get_out_date(file_name):
    out_date_tuple = os.path.getmtime(file_name)
    out_date_tuple = tuple(time.localtime(out_date_tuple))
    out_date = f"{out_date_tuple[0]}年{out_date_tuple[1]:02.0f}月{out_date_tuple[2]:02.0f}日" \
               f"{out_date_tuple[3]:02.0f}时{out_date_tuple[4]:2.0f}分"
    return out_date


# 邮件处理对象，用于模拟登录邮箱和处理邮件
class EmailBox:
    def __init__(self, user, password, from_name='理学院青年大学习学习助手', email_add='smtp.qq.com', port=465):
        self.my_email = smtplib.SMTP_SSL(email_add, port)
        self.my_email.login(user, password)
        self.user = user
        self.from_name = from_name

    def sent(self, to_who, to_add, subject, text, file_path_list=[]) -> None:
        """向指定用户发送指定内容的邮件（可以包含附件）

        :param to_who: 接受邮件用户的称呼
        :param to_add: 接受邮件用户的邮箱
        :param subject: 发送邮件的主题
        :param text: 发送邮件的文本（请使用HTML编写）
        :param file_path_list: 要发送的附件路径的列表（无附件时可省略）
        :return: None

        Example:
        email.sent('boss', 'abcd@outlook.com', '0 . 0', '123',
                ['./Original Study Records/table_1.csv'])
        email.sent('boss', 'abcd@outlook.com', '0 . 0', '123')
        """
        # TODO my_errors.TheTypeError使用过于繁琐，记得修改
        if not type(file_path_list) == type(list()):
            raise my_errors.TheTypeError(type(list()))
        msg = MIMEMultipart()  # 创建要发送的邮件
        msg_text = MIMEText(text, 'html', 'utf-8')
        msg.attach(msg_text)  # 给要发送的邮件添加文本
        for file_path in file_path_list:
            file_name = os.path.basename(file_path)
            input('next')
            print(file_name)
            input()
            with open(file_path, 'rb') as f:
                file_part = MIMEApplication(f.read())  # 读取文件
            # 给文件添加识别标志，前两个保证文件能被浏览和下载，第三个用于给文件命名
            file_part.add_header('Content-Disposition', 'attachment', filename=file_name)
            # file_part.add_header('Content-Disposition', 'attachment', file_name)
            msg.attach(file_part)

        msg['subject'] = subject  # 编辑邮件主题
        msg['from'] = self.from_name  # 编辑发件人名称
        msg['to'] = to_who  # 编辑收件人名称
        self.my_email.sendmail(self.user, to_add, msg.as_string())  # 发送


