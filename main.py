import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename
from time import sleep
import xlrd
import pyperclip
import threading
import zmail
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from tkinter import font
from collections import OrderedDict

spacing = 50


def send(from_addr, passwd, to_addr, subject, content, title):
    server = smtplib.SMTP()
    server.connect('mail.fudan.edu.cn', 25)
    server.login(from_addr, passwd)

    msg = MIMEText(content, 'plain', 'utf-8')
    msg['From'] = Header(title, 'utf-8')
    msg['From'].append(' <' + from_addr + '>', 'ascii')
    msg['To'] = Header(to_addr)
    msg['Subject'] = Header(subject, 'utf-8')

    server.sendmail(from_addr, to_addr, msg.as_string())

    server.quit()


def thread_it(func, *args):
    t = threading.Thread(target=func, args=args)
    t.setDaemon(True)
    t.start()


def entry_text(win, x_l, x_e, line, text, show=None, width=15, y=None):
    if y is None:
        y = line * spacing + 10

    l = tk.Label(win, text=text).place(x=x_l, y=y)
    e = tk.Entry(win, show=show, width=width)
    e.place(x=x_e, y=y)

    return e


def entry_var(win, x_l, x_e, line, text, show=None, width=15, y=None):
    if y is None:
        y = line * spacing + 10

    l = tk.Label(win, text=text).place(x=x_l, y=y)
    var = tk.StringVar()
    e = tk.Entry(win, textvariable=var, show=show, width=width)
    e.place(x=x_e, y=y)

    return e, var


def button(win, x, line, text, command, y=None, x_l=None, y_l=None):
    if y is None:
        y = line * spacing + 10

    var = tk.StringVar()
    l = tk.Label(win, textvariable=var, fg='red').place(x=x if x_l is None else x_l, y=y + 26 if y_l is None else y_l)

    b = tk.Button(win, text=text, command=command).place(x=x, y=y - 4)

    return var


def to_date(x):
    date_tuple = xlrd.xldate_as_tuple(x.value, 0)
    date = '{:0>4d}'.format(date_tuple[0]) + '年' + '{:0>2d}'.format(date_tuple[1]) + '月' + '{:0>2d}'.format(date_tuple[2]) + '日' if date_tuple[0] != 0 and date_tuple[1] != 0 and date_tuple[2] != 0 else ''
    time = '{:0>2d}'.format(date_tuple[3]) + ':' + '{:0>2d}'.format(date_tuple[4]) if date_tuple[3] != 0 and date_tuple[4] != 0 else ''

    return date + time

def to_str(x):
    if isinstance(x.value, float) and int(x.value) == x.value:
        x.value = int(x.value)

    return str(x.value)

def to_dec(c):
    return ord(c.upper()) - ord('A')


def main():
    global index, content, subject, mail_from, passwd, title, wb, addrs, infos, interval
    addrs = []
    infos = OrderedDict()
    f = open('./subject.txt', 'r', encoding='utf-8')
    subject = f.read()
    f.close()
    f = open('./content.txt', 'r', encoding='utf-8')
    content = f.read()
    f.close()
    f = open('./interval.txt', 'r')
    interval = f.read().strip()
    interval = int(interval) if len(interval) else 1
    f.close()

    win = tk.Tk()
    win.title('通知邮件群发助手')
    win.geometry('1160x620')

    def account_prt():
        global mail_from, passwd, title

        f = open('./login.txt', 'r', encoding='utf-8')
        lines = f.readlines()
        f.close()

        if len(lines) == 0 or len(lines[0].strip()) == 0 or lines[0].strip().isspace():
            var1.set('无登录信息')
        else:
            mail_from = lines[0].strip() + '@fudan.edu.cn'
            passwd = lines[1].strip()
            title = lines[2].strip()
            var1.set(title + ' <' + mail_from + '>')

    def login():
        top = tk.Toplevel()
        top.title('登陆')
        top.geometry('320x200')

        def get_account():
            mail_from = e1.get().strip()
            passwd = e2.get()
            title = e3.get().strip()

            if len(mail_from) == 0 or len(passwd) == 0 or len(title) == 0:
                var2.set('内容不能为空')
                return
            else:
                server = zmail.server(mail_from + '@fudan.edu.cn', passwd, smtp_host='mail.fudan.edu.cn', pop_host='mail.fudan.edu.cn')
                if server.smtp_able() == False:
                    var2.set('账号或密码错误，请重试')
                    return

            f = open('./login.txt', 'w')
            f.writelines([mail_from + '\n', passwd + '\n', title])
            f.close()

            account_prt()

            top.destroy()

        e1 = entry_text(top, 0, 70, 0, '邮箱账号：', width=15)
        tk.Label(top, text='@fudan.edu.cn').place(x=180, y=0 * spacing + 10)
        e2 = entry_text(top, 0, 70, 1, '邮箱密码：', width=25, show='*')
        e3 = entry_text(top, 0, 70, 2, '抬   头：', width=10)

        tk.Button(top, text='登陆', command=get_account).place(x=150, y=3 * spacing + 10 - 4)
        var2 = tk.StringVar()
        tk.Label(top, textvariable=var2, fg='red').place(x=140, y=3 * spacing + 10 + 26)


    def get_wb():
        global wb

        path = askopenfilename()
        var3.set(path)

        try:
            wb = xlrd.open_workbook(path).sheet_by_index(0)
            var4.set('                 加载成功')

        except:
            var4.set('请选取正确的.xls或.xlsx文件')

    def get_info():
        global addrs, infos, wb
        addrs = []
        infos = OrderedDict()

        addr_index = e4.get()
        info_indexes = e5.get()

        f = open('./index.txt', 'w')
        f.writelines([addr_index + '\n', info_indexes])
        f.close()

        if len(addr_index.strip()) == 0:
            var7.set('                      收件地址列号不可为空')
            return

        addr_index = to_dec(addr_index)
        info_indexes = list(filter(lambda x:x.isalpha(), info_indexes.split(' ')))

        for index_ in info_indexes:
            if len(index_) != 1:
                var7.set('                      列号请使用空格分隔')
                return

        info_indexes = list(map(to_dec, info_indexes))

        try:
            addrs = list(map(to_str, wb.col(addr_index)))
            addrs.append(mail_from)

            for i, index_ in enumerate(info_indexes):
                infos[str(i)] = list(map(to_date if wb.cell_type(0, index_) == 3 else to_str, wb.col(index_)))
                infos[str(i)].append('测试信息#' + chr(index_ + ord('A')))
            var7.set('载入成功，共' + str(len(infos)) + '类信息待嵌入' + str(len(addrs)) + '封邮件')

        except:
            var7.set('                          列号输入有误')

    def copy_brace():
        pyperclip.copy('{}')
        var8.set('复制成功')


    def save_mail():
        global subject, content

        subject = t1.get(0.0, 'end').strip()
        content = t2.get(0.0, 'end').strip()

        f = open('./subject.txt', 'w', encoding='utf-8')
        f.write(subject)
        f.close()
        f = open('./content.txt', 'w', encoding='utf-8')
        f.write(content)
        f.close()

        var9.set('保存成功')

    def work(mail_from, passwd, addrs, infos, subject, text, title, interval):

        for i in range(len(addrs)):
            var12.set('正在发送（' + str(i + 1) + '/' + str(len(addrs)) + ')')
            send(mail_from, passwd, addrs[i], subject, text.format(*list(val[i] for val in infos.values())), title)
            if i + 1 < len(addrs):
                sleep(interval)
            else:
                var12.set('发送成功，请在发件箱中查验')

    def get_interval():
        global interval
        interval = int(e6.get().strip())
        f = open('./interval.txt', 'w')
        f.write(str(interval))
        f.close()
        var11.set('保存成功')

    def iterate():
        try:
            thread_it(work, mail_from, passwd, addrs, infos, subject, content, title, interval)
        except:
            var12.set('发送失败')

    def preview():
        global index
        index = 0
        try:
            content.format(*list(val[index] for val in infos.values()))
        except:
            var9.set('正文嵌入错误')
            return
        try:
            addrs[index]
        except:
            var9.set('邮件列号未载入')
            return

        msg_ = MIMEText(content, 'plain', 'utf-8')
        msg_['From'] = Header(title, 'utf-8')
        msg_['From'].append(' <' + mail_from + '>', 'ascii')
        msg_['Subject'] = Header(subject, 'utf-8')

        def pageup():
            global index
            if index - 1 >= 0:
                index -= 1
                var_addr_to.set('收件人： ' + addrs[index])
                var_content.set(content.format(*list(val[index] for val in infos.values())))
                var_index.set(str(index + 1) + '/' + str(len(addrs)))

        def pagedown():
            global index
            if index + 1 < len(addrs):
                index += 1
                var_addr_to.set('收件人： ' + addrs[index])
                var_content.set(content.format(*list(val[index] for val in infos.values())))
                var_index.set(str(index + 1) + '/' + str(len(addrs)))

        top = tk.Toplevel()
        top.title('邮件预览')
        top.geometry('650x680')

        font_subject = font.Font(size=12, weight=font.BOLD)
        font_ = font.Font(size=12)

        var_subject = tk.StringVar()
        var_subject.set(str(msg_['Subject']))
        tk.Message(top, textvariable=var_subject, width=500, font=font_subject).place(x=10, y=0 * spacing + 10)

        var_addr_from = tk.StringVar()
        var_addr_from.set('发件人： ' + str(msg_['From']))
        tk.Message(top, textvariable=var_addr_from, width=500, font=font_).place(x=10, y=0.5 * spacing + 10)

        var_addr_to = tk.StringVar()
        var_addr_to.set('收件人： ' + addrs[index])
        tk.Message(top, textvariable=var_addr_to, width=500, font=font_).place(x=10, y=1 * spacing + 10)

        var_content = tk.StringVar()
        var_content.set(content.format(*list(val[index] for val in infos.values())))
        tk.Message(top, textvariable=var_content, width=600, font=font_).place(x=10, y=2 * spacing + 10)

        tk.Button(top, text='上一封', command=pageup).place(x=230, y=12.5 * spacing + 10)

        var_index = tk.StringVar()
        var_index.set(str(index + 1) + '/' + str(len(addrs)))
        tk.Message(top, textvariable=var_index, width=50).place(x=285, y=12.5 * spacing + 10)

        tk.Button(top, text='下一封', command=pagedown).place(x=345, y=12.5 * spacing + 10)

    tk.Label(win, text='当前登录账号：').place(x=0, y=0 * spacing + 10)
    var1 = tk.StringVar()
    tk.Label(win, textvariable=var1).place(x=85, y=0 * spacing + 10)
    account_prt()

    tk.Button(win, text='切换账户', command=login).place(x=305, y=0 * spacing + 10 - 4)

    e3, var3 = entry_var(win, 0, 120, 1, 'Excel文件所在路径：', width=65)

    button(win, 590, 1, '选择文件', get_wb, x_l=525)
    var4 = button(win, 590, 1, '选择文件', get_wb, x_l=525)

    e4, var5 = entry_var(win, 0, 115, 2, text='邮箱地址所在列号：', width=5)
    e5, var6 = entry_var(win, 210, 300, 2, text='信息所在列号: ', width=15)
    f = open('./index.txt', 'r')
    lines = f.readlines()
    f.close()
    var5.set('' if len(lines) == 0 else lines[0].strip())
    var6.set('' if len(lines) < 2 else lines[1].strip())

    var7 = tk.StringVar()
    tk.Button(win, text='保存并载入信息', command=get_info).place(x=420, y=2 * spacing + 10 - 4)
    tk.Label(win, textvariable=var7, fg='red').place(x=320, y=2 * spacing + 10 + 26)

    tk.Button(win, text='点击复制"{}"', command=copy_brace).place(x=585, y=2 * spacing + 10 - 4)
    var8 = tk.StringVar()
    tk.Label(win, textvariable = var8, fg='red').place(x=595, y=2 * spacing + 10 + 26)

    tk.Label(win, text='主题：').place(x=0, y=3 * spacing + 10)
    t1 = tk.Text(win, height=2, width=50)
    t1.insert(0.0, subject)
    t1.place(x=40, y=3 * spacing + 10)

    tk.Label(win, text='正文：').place(x=0, y=4 * spacing + 10)
    t2 = ScrolledText(win, height=27, width=85)
    t2.place(x=40, y=4 * spacing + 10)
    t2.insert(0.0, content)

    var9 = button(win, 450, 11.5, '保存正文及主题', save_mail, x_l=598, y_l=11.5 * spacing + 10)


    e6, var10 = entry_var(win, 385, 475, 0, '设定发送间隔：')
    var10.set(str(interval))
    tk.Label(win, text='秒/封').place(x=555, y=0 * spacing + 10)
    var11 = tk.StringVar()
    tk.Label(win, textvariable=var11, fg='red').place(x=595, y=0 * spacing + 10 + 26)
    tk.Button(win, text='保存', command=get_interval).place(x=605, y=0 * spacing + 10 - 4)

    button(win, 560, 11.5, '预览', preview)

    var12 = button(win, 38, 11.5, '发送', iterate, y_l=11.5 * spacing + 10, x_l=80)

    tk.Label(win, text='使用帮助：').place(x=685, y=10)

    help = '''———————————————————————————————
准备工作：

登录复旦邮箱
->点击页面右上角【设置】并选择【邮箱设置】
->在"基本信息"条目中选择【参数设置】
->在"发信/写信"条目中，开启【SMTP发信后保存到[已发送]】选项

* 注意：“content/txt”等文本文件请勿更改或删除，并保持与”通   知邮件群发助手.exe“在同一目录下。
———————————————————————————————
Step1：

核对当前登录账号是否有误；
若需更换账户，点击“切换账户”进行账户登录。

———————————————————————————————
Step2：

设定发送时间间隔。

———————————————————————————————
Step3：

点击“选择文件”选择关键信息所在Excel文件，并点击确定。

———————————————————————————————
Step4：

填写Excel文件中，“邮箱地址所在列号”及“信息所在列号”，并点击“保存并载入”，出样“载入成功”字样后继续。

* 注意：列号为大写字母，A为第一列
* 注意：“信息所在列号”条目中需填写所有需要嵌入信息的列号，列  号依照  嵌入顺序填写，以空格分隔。
* 注意：点击“复制{}”将括号复制进剪贴板，以嵌入正文。
———————————————————————————————
Step5：

编辑邮件主题及正文，并点击“保存正文及主题”后继续。

* 注意：请核对嵌入括号数与嵌入信息数相对应。
———————————————————————————————
Step6：

点击发送，等待发送结束后在发件箱中查验。
———————————————————————————————'''

    t3 = ScrolledText(win, height=43, width=62)
    t3.place(x=685, y=40)
    t3.insert(0.0, help)

    win.mainloop()


if __name__ == '__main__':
    index = 0
    main()