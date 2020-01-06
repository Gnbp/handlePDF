from tkinter.filedialog import askopenfilename
from tkinter import messagebox, StringVar, Tk
from tkinter import Label, Entry, Button, Listbox, Radiobutton, Frame, Scrollbar
from tkinter import BOTH, X, Y, LEFT, RIGHT, N, S, E, W, YES
import os
import shutil

from PyPDF2 import PdfFileReader as pdfreader, PdfFileWriter as pdfwriter
import webbrowser
import time



"""
TK.wm_minsize(1440, 776)                  # 设置窗口最小化大小
TK.wm_maxsize(1440, 2800)                 # 设置窗口最大化大小
TK.resizable(width=False, height=True)    # 设置窗口宽度不可变，高度可变
"""


def get_dirpath(filepath):
    return os.path.split(filepath)[0]

def get_filename(filepath):
    return os.path.split(filepath)[-1].split('.')[0]

origin_filepath = ''
origin_filename = ''
new_filepath = ''
check_dirpath = ''


warning_title = '警告'
reminder_title = '提示'


chose_page_msg = '请先选择页码'
del_dir_msg = '缓存文件已清空'
rotate_finish_msg = '翻转完成'
copy_finish_msg = '副本生成完成！'
safe_msg = '为了安全，请先生成副本文件!'

chose_error_msg = '请选择起始文件！'
range_error_msg = '请输入范围，使用‘-’相连，或者输入多个不连续的页数，使用‘+’相连'


# Web open 
WEB_TITLE = r'file://'


# MyTk setting
TITLE = '文件处理'
SIZE = '500x370'


tk = Tk()
tk.title(TITLE)
tk.geometry(SIZE)
tk.columnconfigure(0, weight=1)
origin_path = StringVar()
chose_handler_str = StringVar()
lb_list = StringVar()
radio_var = StringVar()
radio_var.set('90')


ftop = Frame(tk)
ftop.grid(row=0, column=0)

fleft = Frame(tk)
fleft.grid(row=1, column=0, sticky=N+S+W+E) 

fright = Frame(tk)   
fright.grid(row=1, column=1)

fbot = Frame(tk)
fbot.grid(row=2, column=0, sticky=S)


def alerm_msg(tit, msg):
    messagebox.showinfo(title=tit, message=msg)

def chose_file1():
    _path1 = askopenfilename()
    origin_path.set(_path1)

def check_new_path():
    if origin_path.get() != '':
        origin_filepath = origin_path.get()
        base_dir_path = get_dirpath(origin_filepath)
        origin_filename = get_filename(origin_filepath)
        filedirpath = os.path.join(base_dir_path, origin_filename)
        check_newpath = os.path.join(filedirpath, origin_filename+'{}'.format(origin_filepath[-4:]))
        if os.path.exists(check_newpath):
            global new_filepath
            new_filepath = check_newpath

def copy_origin_file():
    if origin_path.get() == '':
        alerm_msg(warning_title, chose_error_msg)
    else:
        global origin_filepath, origin_filename, new_filepath
        origin_filepath = origin_path.get()
        base_dir_path = get_dirpath(origin_filepath)
        origin_filename = get_filename(origin_filepath)
        filedirpath = os.path.join(base_dir_path, origin_filename)
        if not os.path.exists(filedirpath):
            os.mkdir(filedirpath)
        new_filepath = os.path.join(filedirpath, origin_filename+'{}'.format(origin_filepath[-4:]))
        shutil.copyfile(origin_filepath, new_filepath)
        alerm_msg(reminder_title, copy_finish_msg) 
        
def rotate_page(rotate_pages, rotate_direction, file_path):
    if rotate_direction == '90':
        rotate_degree = 90
    elif rotate_direction == '270':
        rotate_degree = 270
    elif rotate_direction == '180':
        rotate_degree = 180
    else:
        print('请输入规定内的翻转角度')
    
    reader_obj = pdfreader(file_path)
    writer_obj = pdfwriter()
    total_pages = reader_obj.getNumPages()
    
    for rpage in range(total_pages):
        if rpage+1 in rotate_pages:
            page_num_obj = reader_obj.getPage(rpage).rotateClockwise(rotate_degree)
        else:
            page_num_obj = reader_obj.getPage(rpage)
        writer_obj.addPage(page_num_obj)
    with open(file_path, 'wb') as fw:
        writer_obj.write(fw)
        

def display_pages(lb_lists=[]):
    check_new_path()
    if new_filepath != '':
        lb_list.set('')
        input_num_str = chose_handler_str.get()
        page_lists = []
        try:
            handler_page = int(input_num_str)
        except ValueError:
            if '-' in input_num_str:
                try:
                    start_num = int(input_num_str.split('-')[0])
                    end_num = int(input_num_str.split('-')[1]) + 1
                except ValueError:
                    alerm_msg(warning_title, range_error_msg)
                    chose_handler_str.set('')
                else:
                    for i in range(start_num, end_num):
                        page_lists.append(i)
            elif '+' in input_num_str:
                page_split_lists = input_num_str.split('+')
                for i in set(page_split_lists):
                    try:
                        page_lists.append(int(i))
                    except ValueError:
                        alerm_msg(warning_title, range_error_msg)
                        chose_handler_str.set('')
                        break
                page_lists.sort()            
            else:
                alerm_msg(warning_title, range_error_msg)
                chose_handler_str.set('')
        else:
            page_lists.append(handler_page)
        finally:
            if page_lists != []:
                chose_handler_str.set('')
                if lb_lists == []:
                    for p in page_lists:
                        lb.insert('end', p)
                    # 设置self.lb的纵向滚动影响scroll滚动条
                    lb.configure(yscrollcommand=scroll.set)
                elif lb_lists == 'allfile':
                    rotate_page(page_lists, radio_var.get(), new_filepath)
                    alerm_msg(reminder_title, rotate_finish_msg)
    else:
        alerm_msg(reminder_title, safe_msg)
        

def open_page_web():
    if new_filepath == '':
        alerm_msg(reminder_title, safe_msg)
    try:
        page_num = str(lb.get(lb.curselection()))
    except Exception:
        alerm_msg(warning_title, chose_page_msg)
    else:
        global check_dirpath
        check_dirpath = os.path.join(get_dirpath(new_filepath), 'web_check')
        if not os.path.exists(check_dirpath):
            os.mkdir(check_dirpath)
        check_file_path = os.path.join(check_dirpath, page_num+'.pdf')
        pdf_reader = pdfreader(new_filepath)
        pdf_writer = pdfwriter()
        pdf_writer.addPage(pdf_reader.getPage(int(page_num)-1))
        with open(check_file_path, 'wb') as fw:
            pdf_writer.write(fw)
        add_web_prefix_file = WEB_TITLE + check_file_path
        new = 0 
        webbrowser.open(add_web_prefix_file, new=new)
        

def rotate_one_page(r_direction):
    if new_filepath == '':
        alerm_msg(warning_title, safe_msg)
    try:
        page_num = str(lb.get(lb.curselection()))
    except Exception:
        alerm_msg(warning_title, chose_page_msg)
    else:
        rotate_page([int(page_num)], r_direction, new_filepath)
        open_page_web()


def del_dir():
    if os.path.exists(check_dirpath):
        shutil.rmtree(check_dirpath)
        alerm_msg(reminder_title, del_dir_msg)


Label(ftop, text='起始路径').grid(row=0, column=0)
Entry(ftop, textvariable=origin_path).grid(row=0, column=1)
Button(ftop, text='路径选择', command=chose_file1).grid(row=0, column=2)
Button(ftop, text='生成副本', command=copy_origin_file).grid(row=0, column=3) 
Label(ftop, text='翻转页码').grid(row=1, column=0)
Entry(ftop, textvariable=chose_handler_str).grid(row=1, column=1)
Button(ftop, text='翻转页数', command=display_pages).grid(row=1, column=2)

r1 = Radiobutton(ftop, text='向左翻九十', variable=radio_var, value='270')
r1.grid(row=2, column=0)
r2 = Radiobutton(ftop, text='向右翻九十', variable=radio_var, value='90')
r2.grid(row=2, column=1)
r2 = Radiobutton(ftop, text='翻一百八', variable=radio_var, value='180')
r2.grid(row=2, column=2)
Button(ftop, text='一键翻转', command=lambda : display_pages('allfile')).grid(row=2, column=4)

lb = Listbox(fleft, listvariable=lb_list)
lb.pack(side=LEFT, fill=BOTH, expand=YES)


scroll = Scrollbar(fleft, command=lb.yview)
scroll.pack(side=RIGHT, fill=Y, pady=10)


b0 = Button(fright, text='浏览器查看', command=open_page_web)
b0.grid(row=0, column=4)

b3 = Button(fright, text='向左翻九十', command=lambda : rotate_one_page('270'))
b3.grid(row=3, column=4)   

b4 = Button(fright, text='向右翻九十', command=lambda : rotate_one_page('90'))
b4.grid(row=4, column=4)

b5 = Button(fright, text='翻一百八', command=lambda : rotate_one_page('180'))
b5.grid(row=5, column=4)

Button(fbot, text='清除查看文件', command=del_dir).pack()
Button(fbot, text='退       出', command=tk.quit).pack()


tk.mainloop()

# if new_filepath != '':
#     finish_time_str = time.strftime('_%m_%d-%H', time.localtime(time.time()))
#     last_filepath = new_filepath[:-4] + finish_time_str + new_filepath[-4:]
#     shutil.copyfile(new_filepath, last_filepath)
#     print('请放心，修改已经备份')

