import json
import os.path
import tkinter
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import *
from tkinter.ttk import Combobox
from typing import Type
import log
from log import log, logQue, levelQue, getTime
import windnd
import xlrd


class ExcelExport(Tk):
    # excel文件夹
    excel_dir = ""
    # excel文件列表
    excel_file_name_list = []
    # json存放路径
    json_dir = ""
    # 0 正常 1 error
    status = 0
    # 0 单文件模式 1 文件夹模式 2 多文件模式
    model = 0

    def __init__(self):
        super().__init__()
        # 初始化主窗口
        self._set_window_()
        # 创建菜单栏
        self._create_menu_bar_()
        # 创建窗口主体
        self._create_body_()
        # 创建右键菜单
        self._create_right_popup_menu()

    def _set_window_(self):
        """
        创建主窗口
        :return:
        """
        log.debug("创建主窗口")
        self.title("转换工具-梦游勇者工作室自用")
        scn_width, scn_height = self.maxsize()
        wm_val = '750x450+%d+%d' % ((scn_width - 750) / 2, (scn_height - 450) / 2)
        self.geometry(wm_val)
        self.protocol('WM_DELETE_WINDOW', self.exit_util)
        log.debug("主窗口创建完毕")

    # 创建整个菜单栏
    def _create_menu_bar_(self):
        """
        创建界面上方菜单栏
        :return:
        """
        log.debug("创建菜单栏")
        menu_bar = Menu(self)
        # 创建菜单
        file_menu = Menu(menu_bar, tearoff=0)
        model_men = Menu(menu_bar, tearoff=0)
        model_men.add_radiobutton(label="单文件模式", command=lambda: self.model_change(0))
        model_men.add_radiobutton(label="多文件模式", command=lambda: self.model_change(2))
        model_men.add_radiobutton(label="文件夹模式", command=lambda: self.model_change(1))
        file_menu.add_cascade(label='模式切换', menu=model_men)
        file_menu.add_command(label='清空数据', accelerator='Ctrl+O', command=self.clear_data)
        file_menu.add_command(label='清空日志', accelerator='Ctrl+S', command=self.clear_log)
        file_menu.add_command(label='全部清空', accelerator='Shift+Ctrl+S', command=self.all_clear)
        file_menu.add_separator()
        file_menu.add_command(label='退出', accelerator='Alt+F4', command=self.exit_util)

        # 在菜单栏上添加菜单标签，并将该标签与相应的联级菜单关联起来
        menu_bar.add_cascade(label='文件', menu=file_menu)

        about_menu = Menu(menu_bar, tearoff=0)
        about_menu.add_command(label='关于', command=lambda: self.show_messagebox('关于'))
        about_menu.add_command(label='帮助', command=lambda: self.show_messagebox('帮助'))
        menu_bar.add_cascade(label='关于', menu=about_menu)
        self["menu"] = menu_bar
        log.debug("菜单栏创建完毕")

    # 创建程序主体
    def _create_body_(self):
        """
        从上到下依次创建窗口主要部件
        :return:
        """
        log.debug("创建窗口主体")

        log.debug("加载选择Excel部分")
        self.excel_lf = LabelFrame(self, text='选择Excel文件')
        self.excel_lf.place(relx=0.05, rely=0.025, relwidth=0.9, relheight=0.125)
        self.excelTip = Label(self.excel_lf, fg="black", font=('宋体', 10), text="Excel文件:")
        self.excelTip.place(relx=0.05, rely=0.05, relheight=0.8, relwidth=0.17)
        self.excel_loc = StringVar(value="点击右侧或拖拽选择文件")
        locEnt = Entry(self.excel_lf, textvariable=self.excel_loc, font=('宋体', 10), state='readonly')
        locEnt.place(relx=0.22, rely=0.05, relwidth=0.6, relheight=0.8)
        windnd.hook_dropfiles(self.excel_lf, func=self.drag_excel_files)
        self.excelBut = Button(self.excel_lf, text="选择文件", font=('宋体', 10), command=self.choose_excel)
        self.excelBut.place(relx=0.83, rely=0.05, relheight=0.8, relwidth=0.15)
        log.debug("加载选择Excel部分完成")

        log.debug("加载Json路径部分")
        self.json_lf = LabelFrame(self, text='选择Json路径')
        self.json_lf.place(relx=0.05, rely=0.175, relwidth=0.9, relheight=0.125)
        excelTip = Label(self.json_lf, fg="black", font=('宋体', 10), text="Json路径:")
        excelTip.place(relx=0.05, rely=0.05, relheight=0.8, relwidth=0.15)
        self.json_dir_loc = StringVar(value="点击右侧或拖拽选择文件夹")
        windnd.hook_dropfiles(self.json_lf, func=self.drag_json_dir)
        locEnt = Entry(self.json_lf, textvariable=self.json_dir_loc, font=('宋体', 10), state='readonly')
        locEnt.place(relx=0.22, rely=0.05, relwidth=0.6, relheight=0.8)
        excelBut = Button(self.json_lf, text="选择文件夹", font=('宋体', 10), command=self.choose_json_dir)
        excelBut.place(relx=0.83, rely=0.05, relheight=0.8, relwidth=0.15)
        log.debug("加载Json路径部分完成")

        log.debug("加载参数设置")
        config_lf = LabelFrame(self, text='参数设置')
        config_lf.place(relx=0.05, rely=0.325, relwidth=0.9, relheight=0.25)

        type_col_tip = Label(config_lf, fg="black", font=('宋体', 10), text="类型所在行:")
        self.type_col = IntVar(value=1)
        type_col = Combobox(config_lf, width=12, font=('宋体', 10), textvariable=self.type_col)
        type_col['values'] = ([x for x in range(1, 20)])
        type_col_tip.place(relx=0.05, rely=0.00, relwidth=0.17, relheight=0.40)
        type_col.place(relx=0.21, rely=0.00, relwidth=0.17, relheight=0.35)

        sheet_col_tip = Label(config_lf, fg="black", font=('宋体', 10), text="选择表单:")
        self.sheet_col = IntVar(value=1)
        sheet_col = Combobox(config_lf, width=12, font=('宋体', 10), textvariable=self.sheet_col)
        sheet_col['values'] = ([x for x in range(1, 20)])
        sheet_col_tip.place(relx=0.05, rely=0.45, relwidth=0.17, relheight=0.40)
        sheet_col.place(relx=0.21, rely=0.45, relwidth=0.17, relheight=0.35)

        name_col_tip = Label(config_lf, fg="black", font=('宋体', 10), text="名称所在行:")
        self.name_col = IntVar(value=2)
        name_col = Combobox(config_lf, width=12, font=('宋体', 10), textvariable=self.name_col)
        name_col['values'] = ([x for x in range(1, 20)])
        name_col_tip.place(relx=0.40, rely=0.0, relwidth=0.17, relheight=0.40)
        name_col.place(relx=0.56, rely=0.0, relwidth=0.17, relheight=0.35)

        excelBut = Button(config_lf, text="开始转换", font=('宋体', 10), command=self.exchange)
        excelBut.place(relx=0.83, rely=0.05, relheight=0.8, relwidth=0.15)
        log.debug("加载参数设置完成")

        log.debug("加载日志模块")
        config_lf = LabelFrame(self, text='日志显示')
        config_lf.place(relx=0.05, rely=0.6, relwidth=0.9, relheight=0.30)
        self.config_text = Text(config_lf, state='disabled')
        scroll = Scrollbar(self.config_text)
        scroll.config(command=self.config_text.yview)
        self.config_text.tag_config('debug', foreground='gray')
        self.config_text.tag_config('info', foreground='green')
        self.config_text.tag_config('warning', foreground='orange')
        self.config_text.tag_config('error', foreground='red')
        self.config_text.tag_config('critical', foreground='red', font=("宋体", 12, "bold"))
        self.config_text.place(relx=0.1, rely=0.1, relwidth=0.8, relheight=0.8)
        self.config_text.config(yscrollcommand=scroll.set)
        scroll.place(relx=0.96, relwidth=0.05, relheight=1)
        self.config_text.after(100, self.update_log)
        log.debug("加载日志模块部分完成")

        log.debug("窗口主题加载完成")
        pass

    # 鼠标右键弹出菜单
    def _create_right_popup_menu(self):
        """
        右键菜单创建,便携操作
        :return:
        """
        popup_menu = Menu(self, tearoff=0)
        model_men = Menu(self, tearoff=0)
        model_men.add_radiobutton(label="单文件模式", command=lambda: self.model_change(0))
        model_men.add_radiobutton(label="多文件模式", command=lambda: self.model_change(2))
        model_men.add_radiobutton(label="文件夹模式", command=lambda: self.model_change(1))
        popup_menu.add_cascade(label='模式切换', menu=model_men)
        popup_menu.add_command(label='清空数据', compound='left', command=self.clear_data)
        popup_menu.add_command(label='清空日志', compound='left', command=self.clear_log)
        popup_menu.add_command(label='全部清空', compound='left', command=self.all_clear)
        popup_menu.add_separator()
        popup_menu.add_command(label='关于', command=lambda: self.show_messagebox("关于"))
        popup_menu.add_command(label='使用说明', command=self.help_window)
        self.bind('<Button-3>', lambda event: popup_menu.tk_popup(event.x_root, event.y_root))

    # 点击选择excel文件
    def choose_excel(self):
        if self.model == 0:
            fileLoc = askopenfilename(filetypes=[("Excel文件", "*.xls;*.xlsx")])
            log.info("获取到Excel文件路径：" + fileLoc)
            self.excel_file_name_list = [fileLoc]
        else:
            fileLoc = askopenfilenames(filetypes=[("Excel文件", "*.xls;*.xlsx")])
            for x in fileLoc:
                log.info("获取到Excel文件路径：" + x)
            self.excel_loc.set("获得Excel文件数量：" + str(len(fileLoc)))
            self.excel_file_name_list = fileLoc

    # 点击选择excel文件夹
    def choose_excel_dir(self):
        dir_loc = askdirectory()
        log.info("获取到Excel文件夹路径：" + dir_loc)
        self.excel_dir = dir_loc
        self.excel_loc.set(dir_loc)

    # 拖拽选择excel文件或者文件夹
    def drag_excel_files(self, files):
        if len(files) > 1:
            self.model_change(2)
            self.excel_file_name_list = [x.decode('gbk') for x in files]
            for x in self.excel_file_name_list:
                log.info("获取到Excel文件路径：" + x)
            self.excel_loc.set("获得Excel文件数量：" + str(len(files)))
        else:

            f = files[0].decode('gbk')
            if os.path.isdir(f):
                self.model_change(1)
                self.excel_dir = f
                log.info("获取文件夹路径：" + f)
            else:
                self.model_change(0)
                self.excel_file_name_list = [f]
                log.info("获取到文件路径:" + f)
            self.excel_loc.set(f)

    # 拖拽json文件夹
    def drag_json_dir(self, files):
        if len(files) > 1:
            log.error("不可拖拽多个文件或者文件夹")
        json_dir = files[0].decode('gbk')
        if os.path.isdir(json_dir):
            log.info("获取文件夹路径:" + json_dir)
            self.json_dir = json_dir
            self.json_dir_loc.set(json_dir)
        else:
            log.warning("请拖拽文件夹到这里！")
            self.json_dir_loc.set("请拖拽文件夹到这里")

    # 选择json文件夹
    def choose_json_dir(self):
        fileLoc = askdirectory()
        log.debug("json文件夹路径:" + fileLoc)
        self.json_dir_loc.set(fileLoc)

    # 转换启动器
    def exchange(self):
        if self.model == 1:
            self.excel_file_name_list.clear()
            for home, dirs, files in os.walk(self.excel_dir):
                for file in files:
                    if file.endswith(".xls") or file.endswith(".xlsx"):
                        self.excel_file_name_list.append(os.path.join(home,file))
        for f in self.excel_file_name_list:
            self.start_exchange(f)

    # 更新日志
    def update_log(self):
        while not logQue.empty():
            self.config_text.configure(state='normal')
            level = levelQue.get()
            if level == 'debug':
                self.config_text.insert(END, logQue.get() + "\n", "debug")
            elif level == 'info':
                self.config_text.insert(END, logQue.get() + "\n", "info")
            elif level == 'warning':
                self.config_text.insert(END, logQue.get() + "\n", "warning")
            elif level == 'error':
                self.config_text.insert(END, logQue.get() + "\n", "error")
            elif level == 'critical':
                self.config_text.insert(END, logQue.get() + "\n", "critical")
            self.config_text.configure(state='disabled')
            self.config_text.see(END)
        self.config_text.after(100, self.update_log)

    def show_messagebox(self, type):
        if type == "帮助":
            self.help_window()
        else:
            messagebox.showinfo("关于", "转换工具_V0.1\n梦游勇者工作室内部自用，不可外传")

    # 帮助窗口
    def help_window(self):
        messagebox.showinfo("使用说明", "帮助菜单施工中.....")

    # 模式更换
    def model_change(self, cur):
        if cur == 1:
            self.model = 1
            self.excel_loc.set("点击或者拖拽选择文件夹")
            self.excelBut.config(text="点击选择文件夹", command=self.choose_excel_dir)
            self.excel_lf.config(text="选择excel文件夹")
            self.excelTip.config(text="Excel文件夹:")
        elif cur == 2:
            self.model = 2
            self.excel_loc.set("点击或者拖拽选择文件")
            self.excelBut.config(text="选择多个文件", command=self.choose_excel)
            self.excel_lf.config(text="多个excel文件")
            self.excelTip.config(text="Excel文件数量:")
        else:
            self.model = 3
            self.excel_loc.set("点击或者拖拽选择文件")
            self.excelBut.config(text="点击选择文件", command=self.choose_excel)
            self.excel_lf.config(text="选择excel文件")
            self.excelTip.config(text="Excel文件:")

    # 待施工
    def clear_data(self):
        self.excel_dir = ""
        self.excel_file_name_list = []
        self.json_dir = ""
        self.json_dir_loc.set("点击或拖拽选择文件夹")
        self.model_change(self.model)
        pass

    # 待施工
    def clear_log(self):
        self.config_text.configure(state='normal')
        self.config_text.delete(0.0, tkinter.END)
        self.config_text.configure(state='disabled')
        pass

    def all_clear(self):
        self.clear_data()
        self.clear_log()

    def exit_util(self):
        if messagebox.askokcancel("退出?", "确定退出吗?"):
            self.destroy()

    def start_exchange(self, excel_loc=Type[str]):
        log.info("开始转换...")
        data = self.readExcel(excel_loc)
        data_json = self.costToJson(data)
        self.loadToJson(excel_loc, data_json)

    def readExcel(self, excel=Type[str]):
        log.info("准备读入Excel:" + excel)
        if not excel.endswith(".xls") and not excel.endswith(".xlsx"):
            messagebox.showerror('读取错误', '非Excel文件，检查后重新点击')
            return None

        book = xlrd.open_workbook(excel)
        sheet = book.sheet_by_index(self.sheet_col.get()-1)
        type_dict = {}
        name_dict = {}
        for x in range(0, sheet.ncols):
            tp = sheet.cell(self.type_col.get()-1, x).value
            if str(tp).lower().find('str') != -1:
                tp = 'str'
            elif str(tp).lower().find('int') != -1:
                tp = 'int'
            elif str(tp).lower().find('bool') != -1:
                tp = 'bool'
            else:
                tp = 'float'
            type_dict[x] = tp
        for x in range(0, sheet.ncols):
            name_dict[x] = sheet.cell(self.name_col.get()-1, x).value
        for (k, v) in type_dict.items():
            print(k, v)
        js = []
        for x in range(0, sheet.nrows):
            if x == self.type_col.get()-1 or x == self.name_col.get()-1:
                continue
            cur_d = {}
            for y in range(0,sheet.ncols):
                v = sheet.cell(x,y).value
                if type_dict[y] == 'str':
                    v = str(v)
                elif type_dict[y] == 'int':
                    v = int(v)
                elif type_dict[y] == 'bool':
                    v = bool(v)
                else:
                    v = float(v)
                cur_d[name_dict[y]] = v
            js.append(cur_d)
        return js

    def costToJson(self, data):

        print(data)
        return data

    def loadToJson(self, file_name, data):
        file_name = os.path.basename(file_name).replace(".xlsx", "").replace(".xls", "") + ".json"
        print(file_name)
        if not os.path.isdir(self.json_dir):
            messagebox.showerror('准备错误', '该文件夹路径不存在')
            log.error("Json文件夹路径出现错误")
        files = os.listdir(self.json_dir)
        if file_name in files:
            log.warning("存在同名文件，即将修改旧文件名称")
            old_file = os.path.join(self.json_dir, file_name)
            new_file = file_name.replace(".json", "") + getTime() + ".old"
            new_file = new_file.replace(":", "-")
            new_file = os.path.join(self.json_dir, new_file)
            os.renames(old_file, new_file)
        file_path = os.path.join(self.json_dir, file_name)
        log.info("当前json文件位置为" + file_path)
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        print("-----------------------------")
        log.info(file_name + "文件转换完成")



