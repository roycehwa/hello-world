
"""
function:   IMAP收取邮件
detail:     支持最后一封邮件的文本打印以及所有附件的下载
"""
 
from calendar import c
import email
import email.header
import imaplib
from bs4 import BeautifulSoup
from datetime import datetime
import configparser
import os

 
 
class IMAP:

    def __init__(self, platform, criteria):
        config = configparser.ConfigParser()
        config.read(".\email\config\settings.ini")
        self.user_id = config[platform]['user']
        self.password = config[platform]['key']
        self.imap_server = config[platform]['imap']
        self.conn = None
        self.criteria = criteria
        self.selector =""
 
    def login(self):
        try:
            serv = imaplib.IMAP4_SSL(self.imap_server, 993)  
            print('imap4 服务器连接成功')
        except Exception as e:
            print('imap4 服务器连接失败:', e)
            exit(1)
 
        try:
            serv.login(self.user_id, self.password)
            print('imap4 登录成功')
            self.conn = serv
            
        except Exception as e:
            print('imap4 登录失败：', e)
            exit(1)
 
    def loginout(self):
        """
        登出邮件服务器
        :param conn: imap连接
        """
        self.conn.close
        self.conn.logout()

    def select(self):
        
        UC = self.criteria.pop('unseen')
        if self.criteria.pop('all'):
            self.selector = "(ALL{})"
        else:
            self.selector = " ".join(['{} "{}"'.format(key.upper(), value) for key,value in self.criteria.items()]) + " {}"
            self.selector = '({})'.format(self.selector)
        unseen = " UNSEEN" if UC else ""
        
        self.selector = self.selector.format(unseen)
        print(self.selector)

        self.conn.select()
        ret, data = self.conn.search(None, IMAP.selector)  # 所有邮件
        email_list = data[0].split()
        print(email_list)

        if len(email_list) == 0:
            print('收件箱为空，已退出')
            exit(1)

        else:
            return email_list

    def get_content(self, scope):
        """
        获取指定邮件scope，解析内容
        :param conn: imap连接
        :return:
        """

        box = []

        for item in scope:
            ret, data = self.conn.fetch(item, '(RFC822)')

            try:
                mmsg = email.message_from_string(data[0][1].decode('gbk'))

            except UnicodeDecodeError:
                mmsg = email.message_from_string(data[0][1].decode('utf-8'))

            box.append(self._parse_content(mmsg)) # 提取邮件中文字内容

        return box



    def _parse_content(self, msg):

        content = {}
        body_text = []

        sub = msg.get('subject')
        email_from = msg.get('from')
        email_to = msg.get('to')

        sub_text = email.header.decode_header(sub) # 返回列表，每个元素为一个元祖，元祖中0为编码文字，1为编码方法
        email_from_text = email.header.decode_header(email_from)
        email_to_text = email.header.decode_header(email_to)

        # 如果是特殊字符，元组的第二位会给出编码格式，需要转码
        if sub_text[0]:
            content["Subject"] = self.tuple_to_str(sub_text[0])
        email_from_detail = ''
        for i in range(len(email_from_text)):
            email_from_detail = email_from_detail + self.tuple_to_str(email_from_text[i])
        email_to_detail = ''
        for i in range(len(email_to_text)):
            email_to_detail = email_to_detail + self.tuple_to_str(email_to_text[i])

        content["from"] = email_from_detail
        content["to"] = email_to_detail
 
        # 通过walk可以遍历出所有的内容
        for part in msg.walk():
            # 这里要判断是否是multipart，如果是，数据没用丢弃
            if not part.is_multipart():
    
                # 内容类型
                content_type = part.get_content_type()
                # 如果是附件，这里就会取出附件的文件名，以下两种方式都可以获取
                # name = part.get_param("name")
                name = part.get_filename()

                if name:
                    # 附件中文名获取到的是=?GBK?Q?=D6=D0=CE=C4=C3=FB.docx?=(中文名.docx)格式，需要将其解码为bytes格式
                    trans_name = email.header.decode_header(name)
                    if trans_name[0][1]:
                        # 将bytes格式转为可读格式
                        file_name = trans_name[0][0].decode(trans_name[0][1])
                    else:
                        file_name = trans_name[0][0]
                    print('开始下载附件:', file_name)
                    attach_data = part.get_payload(decode=True)  # 解码出附件数据，然后存储到文件中
                    try:
                        f = open(file_name, 'wb')  # 注意一定要用wb来打开文件，因为附件一般都是二进制文件
                    except Exception as e:
                        print(e)
                        f = open('tmp', 'wb')
                    f.write(attach_data)
                    f.close()
                    print('附件下载成功:', file_name)
                    content["attach"] = file_name

                else:
                    # 文本内容
                    txt = part.get_payload(decode=True)  # 解码文本内容

                    # 分别解释text/html和text/plain两种类型，纯文本解释起来较简单，两种类型内容一致
                    if content_type == 'text/html':
                        print('处理邮件正文(text/html)。。。')
                        # 注意不同邮件平台的HTML结构不同，有一层包裹或者二层包裹
                        try: 
                            soup = BeautifulSoup(str(txt, "gbk"), 'lxml')
                        except UnicodeDecodeError:
                            soup = BeautifulSoup(str(txt, "utf-8"), 'lxml')

                        div_data = soup.find_all('div')
                        if len(div_data) > 1:
                            for each in div_data[1:]:
                                if each.text:
                                    body_text.append(each.text)
                        else:
                            for each in soup.find_all('p'):
                                if each.text:
                                    body_text.append(each.text)

                    elif content_type == 'text/plain':
                        print('以下是邮件正文(text/plain)：')
                        # 纯文本格式为bytes，不同邮件服务器较统一
                        try:
                            body_text.append(str(txt, 'utf-8'))
                        except UnicodeDecodeError:
                            body_text.append(str(txt, 'gbk'))
                    
                    content['body_text'] = body_text
            
        return content
 
    def front(self):
        """
        使用163邮箱，必须在select之前上传客户端身份信息,否则报错
        """
        imaplib.Commands['ID'] = 'AUTH'
        # 如果使用163邮箱，需要上传客户端身份信息
        args = ("name", "18602102347", "contact", "18602102347@163.com", "version", "1.0.0", "vendor", "myclient")
        typ, dat = self.conn._simple_command('ID', '("' + '" "'.join(args) + '")')
      
 
    def tuple_to_str(self, tuple_):
        """
        元组转为字符串输出
        :param tuple_: 转换前的元组，QQ邮箱格式为(b'\xcd\xf5\xd4\xc6', 'gbk')或者(b' <XXXX@163.com>', None)，163邮箱格式为('<XXXX@163.com>', None)
        :return: 转换后的字符串
        """
        if tuple_[1]:
            out_str = tuple_[0].decode(tuple_[1])
        else:
            if isinstance(tuple_[0], bytes):
                out_str = tuple_[0].decode('gbk')
            else:
                out_str = tuple_[0]
        return out_str
 
 
if __name__ == '__main__':
    cr = {"all": True, "unseen": True, "since": "10-May-2021"}
    IMAP = IMAP('163', cr)
    IMAP.login()
    IMAP.front()
    emails = IMAP.select()
    tag = IMAP.get_content(emails)
    for item in tag:
        print(len(item['body_text']))
    IMAP.loginout()
    
    

