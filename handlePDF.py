import os
import shutil
import webbrowser
from threading import Thread
from tkinter.filedialog import askopenfilename
from tkinter import messagebox, StringVar, Tk
from tkinter import Label, Entry, Button, Listbox, Radiobutton, Frame, Scrollbar
from tkinter import BOTH, X, Y, LEFT, RIGHT, N, S, E, W, YES


from PyPDF2 import PdfFileReader as pdfreader, PdfFileWriter as pdfwriter
from PyPDF2.utils import PdfReadError




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
filedirpath = ''
check_newpath = ''
total_page_nums = 0


reminder_title = '提示'
warning_title = '警告'
error_title = '错误'


warning_chose_page_msg = '请先选择页码!'
warning_safe_msg = '为了安全，请先生成副本文件!'
warning_cover_copy_msg = '副本文件已存在，是否覆盖？'
warning_file_msg = '共 0 页代表不能文件不能被操作，请使用Adobe reader打开该文件后另存为新文件，然后操作新文件'

reminder_del_dir_msg = '缓存文件已清空!'
reminder_rotate_finish_msg = '翻转完成!'
reminder_copy_finish_msg = '副本生成完成！'

error_chose_path_msg = '请选择起始文件！'
error_range_msg = '请输入范围，使用‘-’相连，或者输入多个不连续的页数，使用‘+’相连!'
error_file_msg = '文件操作失败，请使用Adobe reader打开该文件后另存为新文件，然后操作新文件'

# Web open 
WEB_TITLE = r'file://'


# MyTk setting
TITLE = '文件处理'
SIZE = '500x370'


tk = Tk()
tk.title(TITLE)
tk.geometry(SIZE)
tk.columnconfigure(0, weight=1)
tk.resizable(0, 0)
origin_path = StringVar()
chose_handler_str = StringVar()
lb_list = StringVar()
radio_var = StringVar()
radio_var.set('90')
pagetextvar = StringVar()
pagetextvar.set('共__页')


ftop = Frame(tk)
ftop.grid(row=0, column=0)

fleft = Frame(tk)
fleft.grid(row=1, column=0, sticky=N+S+W+E) 

fright = Frame(tk)   
fright.grid(row=1, column=1)

fbot = Frame(tk)
fbot.grid(row=2, column=0, sticky=S)


def alerm_msg(tit, msg):
    if tit == '提示':
        messagebox.showinfo(tit, msg)
    elif tit == '警告':
        messagebox.showwarning(tit, msg)
    elif tit == '错误':
        messagebox.showerror(tit, msg)
    

def alerm_chose_msg(tit, msg):
    return messagebox.askokcancel(tit, msg)


def set_pdf_pages(filepath):
    if filepath != '':
        page_nums = '__'
        try:
            page_nums = pdfreader(filepath).getNumPages()
        except PdfReadError:
            page_nums = 0
        finally:
            pagetextvar.set('共 '+str(page_nums)+' 页')
            if page_nums == 0:
                alerm_msg(warning_title, warning_file_msg)
            if page_nums > 0:
                global total_page_nums
                total_page_nums = page_nums
            
        
def chose_file1():
    _path1 = askopenfilename()
    origin_path.set(_path1)
    t1 = Thread(target=set_pdf_pages, args=(_path1,))
    t1.start()
    
    
def check_new_path():
    if origin_path.get() != '':
        global origin_filepath, origin_filename, filedirpath, new_filepath, check_newpath
        origin_filepath = origin_path.get()
        base_dir_path = get_dirpath(origin_filepath)
        origin_filename = get_filename(origin_filepath)
        filedirpath = os.path.join(base_dir_path, origin_filename)
        check_newpath = os.path.join(filedirpath, origin_filename+'{}'.format(origin_filepath[-4:]))
        if os.path.exists(check_newpath):
            new_filepath = check_newpath
            return True
    else:
        alerm_msg(error_title, error_chose_path_msg)
    

def copy_origin_file():
    check_new_path()
    if not os.path.exists(filedirpath) and filedirpath != '':
        os.mkdir(filedirpath)
    if os.path.exists(new_filepath) and new_filepath == check_newpath:
        is_cover = alerm_chose_msg(warning_title, warning_cover_copy_msg)
        if is_cover:
            shutil.copyfile(origin_filepath, new_filepath)
            alerm_msg(reminder_title, reminder_copy_finish_msg)
    elif check_newpath != '':
        shutil.copyfile(origin_filepath, check_newpath)
        alerm_msg(reminder_title, reminder_copy_finish_msg)
            
        
def rotate_page(rotate_pages, rotate_direction, file_path):
    if rotate_direction == '90':
        rotate_degree = 90
    elif rotate_direction == '270':
        rotate_degree = 270
    elif rotate_direction == '180':
        rotate_degree = 180
    else:
        print('请输入规定内的翻转角度')
    try:
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
    except PdfReadError:
        alerm_msg(error_title, error_file_msg)
        return 'file_error'
    else:
        alerm_msg(reminder_title, reminder_rotate_finish_msg)
        

def display_pages(lb_lists=[]):
    flag = check_new_path()
    if flag:
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
                        end_num = int(input_num_str.split('-')[1])
                    except ValueError:
                        alerm_msg(error_title, error_range_msg)
                        chose_handler_str.set('')
                    else:
                        if start_num > end_num:
                            start_num, end_num = end_num, start_num 
                        for i in range(start_num, end_num+1):
                            page_lists.append(i)
                elif '+' in input_num_str:
                    page_split_lists = input_num_str.split('+')
                    for i in set(page_split_lists):
                        try:
                            page_lists.append(int(i))
                        except ValueError:
                            alerm_msg(error_title, error_range_msg)
                            chose_handler_str.set('')
                            break
                    page_lists.sort()            
                else:
                    alerm_msg(error_title, error_range_msg)
                    chose_handler_str.set('')
            else:
                page_lists.append(handler_page)
            finally:
                if page_lists != []:
                    chose_handler_str.set('')
                    if lb_lists == []:
                        for p in page_lists:
                            if p > total_page_nums:
                                break
                            lb.insert('end', p)
                        # 设置self.lb的纵向滚动影响scroll滚动条
                        lb.configure(yscrollcommand=scroll.set)
                    elif lb_lists == 'allfile':
                        t3 = Thread(target=rotate_page, args=(page_lists, radio_var.get(), new_filepath))
                        t3.start() 
    elif origin_filepath !='':
        alerm_msg(warning_title, warning_safe_msg)
        

def open_page_web():
    flag = check_new_path()
    if flag:
        try:
            page_num_str = str(lb.get(lb.curselection()))
        except Exception:
            alerm_msg(warning_title, warning_chose_page_msg)
        else:
            global check_dirpath
            check_dirpath = os.path.join(get_dirpath(new_filepath), 'web_check')
            if not os.path.exists(check_dirpath):
                os.mkdir(check_dirpath)
            check_file_path = os.path.join(check_dirpath, page_num_str+'.pdf')
            try:
                pdf_reader = pdfreader(new_filepath)
                pdf_writer = pdfwriter()
                pdf_writer.addPage(pdf_reader.getPage(int(page_num_str)-1))
                with open(check_file_path, 'wb') as fw:
                    pdf_writer.write(fw)
            except PdfReadError:
                alerm_msg(error_title, error_file_msg)
            else:
                add_web_prefix_file = WEB_TITLE + check_file_path
                new = 0 
                webbrowser.open(add_web_prefix_file, new=new)
    elif origin_filepath !='':
        alerm_msg(warning_title, warning_safe_msg)
        

def rotate_one_page(r_direction):
    flag = check_new_path()
    if flag:
        try:
            page_num = str(lb.get(lb.curselection()))
        except Exception:
            alerm_msg(warning_title, warning_chose_page_msg)
        else:
            file_flag = rotate_page([int(page_num)], r_direction, new_filepath)
            if file_flag != 'file_error':
                threading_open_page()
    elif origin_filepath !='':
        alerm_msg(warning_title, warning_safe_msg)
        

def threading_open_page():
    t2 = Thread(target=open_page_web)
    t2.start()


def del_dir():
    if os.path.exists(check_dirpath):
        shutil.rmtree(check_dirpath)
        alerm_msg(reminder_title, reminder_del_dir_msg)


Label(ftop, text='起始路径').grid(row=0, column=0)
Entry(ftop, textvariable=origin_path).grid(row=0, column=1)
Button(ftop, text='路径选择', command=chose_file1).grid(row=0, column=2)
Button(ftop, text='生成副本', command=copy_origin_file).grid(row=0, column=3) 
Label(ftop, text='翻转页码').grid(row=1, column=0)
Entry(ftop, textvariable=chose_handler_str).grid(row=1, column=1)
Button(ftop, text='翻转页数', command=display_pages).grid(row=1, column=2)
Label(ftop, textvariable=pagetextvar).grid(row=1,column=3)

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

b0 = Button(fright, text='浏览器查看', command=threading_open_page)
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


