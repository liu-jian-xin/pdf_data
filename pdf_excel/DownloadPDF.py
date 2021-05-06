# encoding:utf-8


import os
import re
import logging
import shutil

import poplib, email, telnetlib
import datetime, time, sys, traceback
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from dateutil.relativedelta import relativedelta


#文件重命名
def rename(olddir,dirname,new_name):
    filetype = os.path.splitext(olddir)[1]
    print(filetype)
    newdir = os.path.join(dirname,new_name + filetype)
    os.rename(olddir, newdir)


# 加载配置文件
def loadSettingFile(KEYWORDS_Path):
    logging.info('>>>Loading setting file:%s' % os.path.basename(KEYWORDS_Path))
    PathList = {}  # 储存路径列表
    with open(KEYWORDS_Path, 'r', encoding='UTF-8') as fp:
        lines_kw = fp.readlines( )
        for line in lines_kw:
            # print(line)
            line = line.rstrip('\n')  # 删除行尾的换行符
            if re.match(r'^#', line):  # 注释内容，忽略
                pass
            else:
                Type, Path = line.split('=')  # 获得路径
                PathList[Type] = Path
                logging.info('>>>Content:\n %s' % PathList)
    logging.info('>>>Loading setting file done!')
    return PathList


#创建文件夹，如果文件夹存在就删除
def create_folder(FOLDER_RULE):
    path = os.path.join(os.getcwd(), 'all')
    if os.path.exists(path):
        try:
            shutil.rmtree(path)
        except PermissionError:
            print(f'{path},The file in the directory is open. Please close it and execute again')
            os.system("pause")
            time.sleep(3)
            sys.exit()
    os.makedirs(path)
    for folder in FOLDER_RULE:
        path = os.path.join(os.getcwd(),folder)
        if os.path.exists(path):
            try:
                shutil.rmtree(path)
            except PermissionError:
                print(f'{path},The file in the directory is open. Please close it and execute again')
                os.system("pause")
                time.sleep(3)
                sys.exit()
        os.makedirs(path)


class down_email():

    def __init__(self, user, password, eamil_server):
        # 输入邮件地址, 口令和POP3服务器地址:
        self.user = user
        # 此处密码是授权码,用于登录第三方邮件客户端
        self.password = password
        self.pop3_server = eamil_server

    # 字符编码转换
    # @staticmethod
    def decode_str(self, str_in):
        value, charset = decode_header(str_in)[0]
        if charset:
            value = value.decode(charset)
        return value

    # 解析邮件,获取附件
    def get_att(self, msg_in, pdf_rule,folder_rule,date3,id,Subject):
        attachment_files = []
        for part in msg_in.walk():
            # 获取附件名称类型
            file_name = part.get_param("name")  # 如果是附件，这里就会取出附件的文件名

            if file_name:
                h = email.header.Header(file_name)
                # 对附件名称进行解码
                dh = email.header.decode_header(h)
                filename = dh[0][0]
                if dh[0][1]:
                    # 将附件名称可读化
                    filename = self.decode_str(str(filename, dh[0][1]))
                    # print(filename)
                    # filename = filename.encode("utf-8")
                if (".PDF" in filename or ".pdf" in filename):

                    # 下载附件
                    # print(f'正在读取{filename}')
                    data = part.get_payload(decode=True)
                    filename = filename.replace('/','_')
                    # now = str(int(time.time()))
                    att_file = open(os.path.join(os.getcwd(), 'all',str(id)+'_'+filename), 'wb')
                    att_file.write(data)  # 保存附件
                    att_file.close()
                    print(f'Save【{filename}】to【all】directory')
                    for index in range(len(folder_rule)):
                        if pdf_rule[index] in filename:
                            if (index == 4 or index == 5):
                                if '0005/' in Subject:
                                    att_file = open(os.path.join(os.getcwd(), folder_rule[4], str(id) + '_' + filename),'wb')
                                    att_file.write(data)  # 保存附件
                                    att_file.close()
                                    break
                                elif '0002/' in Subject:
                                    att_file = open(os.path.join(os.getcwd(), folder_rule[5], str(id) + '_' + filename),'wb')
                                    att_file.write(data)  # 保存附件
                                    att_file.close()
                                    break
                            # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
                            att_file = open(os.path.join(os.getcwd(), folder_rule[index],str(id)+'_'+filename), 'wb')
                            att_file.write(data)  # 保存附件
                            att_file.close()
                            print(f'Save【{filename}】to【{folder_rule[index]}】directory')
                            # print(os.path.join(os.getcwd(), folder_rule[index],filename))
                            attachment_files.append(filename)
                            if 'Otto' in pdf_rule[index]:
                                new_file = filename.split('.')[0]
                                rename(os.path.join(os.getcwd(), folder_rule[index],str(id)+'_'+filename), os.path.join(os.getcwd(), folder_rule[index]),new_file+'_'+str(date3))

                else:
                    # print("非PDF文件")
                    continue
            #不是附件
            else:
                continue
                # 不是附件，是文本内容
                # print("不是附件，是文本内容")
                # print(self.get_content(part))
                # # 如果ture的话内容是没用的
                # if not part.is_multipart():
                #     # 解码出文本内容，直接输出来就可以了。
                #     print(part.get_payload(decode=True).decode('utf-8'))

        return attachment_files
    #连接服务器，获取指定日期的邮件
    def run_ing(self,days,pdf_rule,folder_rule):
        end_day = datetime.datetime.now().strftime("%Y-%m-%d")
        str_day = (datetime.datetime.now() - relativedelta(days=days)).strftime("%Y-%m-%d")  # 日期赋值
        print(f'The date of getting PDF mail is {str_day}to {end_day}')
        # print(str_day)
        # str_day = "2021-03-26"
        # 连接到POP3服务器,有些邮箱服务器需要ssl加密，可以使用poplib.POP3_SSL
        try:
            telnetlib.Telnet(self.pop3_server, 995)
            self.server = poplib.POP3_SSL(self.pop3_server, 995, timeout=10)
            # print('try')
        except:
            time.sleep(5)
            self.server = poplib.POP3(self.pop3_server, 110, timeout=10)
            # print('except')
        # 身份认证:
        # print(self.user)
        # print(self.password)
        self.server.user(self.user)
        self.server.pass_(self.password)
        # 返回邮件数量和占用空间:
        print('Number of messages %s. Occupied space: %s' % self.server.stat())
        # list()返回所有邮件的编号:
        resp, mails, octets = self.server.list()
        # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
        # print(mails)
        index = len(mails)
        # print(f'共{index}多封邮件')
        # print()
        for i in range(index, 0, -1):  # 倒序遍历邮件

            # for i in range(1, index + 1):# 顺序遍历邮件
            resp, lines, octets = self.server.retr(i)
            # lines存储了邮件的原始文本的每一行,
            # 邮件的原始文本:
            msg_content = b'\r\n'.join(lines).decode('utf-8')
            # print(msg_content)
            # 解析邮件:
            msg = Parser().parsestr(msg_content)
            # 获取发件人
            From = parseaddr(msg.get('from'))[1]
            # 获取收件人
            To = parseaddr(msg.get('To'))[1]
            # 抄送人
            Cc = parseaddr(msg.get_all('Cc'))[1]
            # 获取邮件主题
            Subject = self.decode_str(msg.get('Subject'))
            # print('from:%s,to:%s,Cc:%s,subject:%s' % (From, To, Cc, Subject))
            # 获取邮件时间,格式化收件时间
            date1 = time.strptime(msg.get("Date")[0:24], '%a, %d %b %Y %H:%M:%S')
            # 邮件时间格式转换
            date2 = time.strftime("%Y-%m-%d", date1)
            date3 = time.strftime("%Y%m%d", date1)
            # 从最近的日期获取邮件
            if date2 < str_day and (index + 1 - i) > 100:
                # attach_file = self.get_att(msg, str_day)
                break  # 倒叙用break
                # continue # 顺叙用continue
            elif date2 <= str_day:
                continue
            else:
                id = (index + 1 - i)
                print(f'Reading {id} message,Read on {date2}')
                # 获取附件
                try :
                # 解析邮件，获取附件
                    attach_file = self.get_att(msg, pdf_rule,folder_rule,date3,id,Subject)
                except Exception as e:
                    print('Unknown exception',e,type(e))
                    time.sleep(2)
                # print(attach_file)

        # 可以根据邮件索引号直接从服务器删除邮件:
        # self.server.dele(7)
        self.server.quit()


if __name__ == '__main__':
    # 加载配置文件
    path_dict = loadSettingFile('./KEYWORDS.txt')
    try:
        # 输入邮件地址, 口令和POP3服务器地址:
        # user = input('请输入邮箱账号：')  # 'dsada.liu@tcl.com'
        # password = input('请输入邮箱密码：')  # 'das@0da00817'
        user = path_dict['user']
        password = path_dict['password']
        eamil_server = path_dict['eamil_server']
        print(user)
        print(password)
        print(eamil_server)
        #匹配规则
        PDF_RULE = path_dict['PDF_RULE'].split(';')
        #目录规则
        FOLDER_RULE = path_dict['FOLDER_RULE'].split(';')
        days = int(path_dict['days'])
        # print('tte\pay.Ee')
        # print('tte\\pay.Ee')
        # 创建文件夹
        create_folder(FOLDER_RULE)
        #创建down_email对象，赋予账号、密码、服务器地址
        email_class = down_email(user=user, password=password, eamil_server=eamil_server)
        # print('email_class')
        #连接服务器，获取指定日期的邮件
        email_class.run_ing(days,PDF_RULE,FOLDER_RULE)
        os.system("pause")
        print('The program is running and exiting normally。。。')
        time.sleep(2)
    except Exception as e:
        import traceback
        ex_msg = '{exception}'.format(exception=traceback.format_exc())
        os.system("pause")
        time.sleep(2)
        print(ex_msg)