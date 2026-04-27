import os
import sys
import shutil
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 新增：用于加载 jpg 格式的图片作为小图标和任务栏略缩图
try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

LANG = {
    "zh": {
        "title": "批量修改文件名工具",
        "lang_btn": "[EN]",
        "input_path": "输入文件路径:",
        "select_path": "[ 选择路径 ]",
        "same_path": "是否输出到同一路径下",
        "output_path": "输出路径:",
        "custom_settings": " 自定义重命名设置 ",
        "add_mod_prefix": "增加/修改前缀",
        "replace_with": "替换为",
        "add_mod_suffix": "增加/修改后缀",
        "seq_naming": "按序号命名",
        "custom_naming": "自定义命名",
        "new_name": "新名称:",
        "export_btn": "[ 导出文件列表 ]",
        "export_format": "输出格式:",
        "include_ext": "包含扩展名",
        "run_btn": ">>> 开始执行重命名 <<<",
        "prefix_modes": ["增加前缀", "修改前缀"],
        "suffix_modes": ["增加后缀", "修改后缀"],
        "filename_header": "文件名",
        "msg_warn_input": "请先选择有效的文件路径！",
        "msg_no_files": "该路径下没有文件。",
        "msg_warn_output": "请选择输出路径！",
        "msg_done": "批量修改完成！\n共成功处理 {} 个文件。",
        "msg_export_succ": "文件列表已成功导出到:\n{}",
        "msg_err": "发生错误: {}",
        "msg_no_openpyxl": "尚未安装 openpyxl 库，无法导出为 xlsx。\n请在终端执行: pip install openpyxl",
        "file_list_name": "文件名列表_导出",
        "error": "错误",
        "warning": "警告",
        "info": "提示",
        "success": "成功",
        "select_dir": "选择目录",
        "watermark": "伍冠宇出品 必属精品"
    },
    "en": {
        "title": "Batch Rename Tool",
        "lang_btn": "[中文]",
        "input_path": "Input Path:",
        "select_path": "[ Select ]",
        "same_path": "Output to same path",
        "output_path": "Output Path:",
        "custom_settings": " Custom Rename Settings ",
        "add_mod_prefix": "Add/Modify Prefix",
        "replace_with": "Replace with",
        "add_mod_suffix": "Add/Modify Suffix",
        "seq_naming": "Sequential Naming",
        "custom_naming": "Custom Naming",
        "new_name": "New Name:",
        "export_btn": "[ Export File List ]",
        "export_format": "Format:",
        "include_ext": "Include Ext",
        "run_btn": ">>> EXECUTE RENAME <<<",
        "prefix_modes": ["Add Prefix", "Modify Prefix"],
        "suffix_modes": ["Add Suffix", "Modify Suffix"],
        "filename_header": "Filename",
        "msg_warn_input": "Please select a valid input path first!",
        "msg_no_files": "No files found in the specified path.",
        "msg_warn_output": "Please select an output path!",
        "msg_done": "Batch rename completed!\nSuccessfully processed {} files.",
        "msg_export_succ": "File list successfully exported to:\n{}",
        "msg_err": "Error occurred: {}",
        "msg_no_openpyxl": "openpyxl is not installed. Cannot export to xlsx.\nRun: pip install openpyxl",
        "file_list_name": "Exported_File_List",
        "error": "Error",
        "warning": "Warning",
        "info": "Info",
        "success": "Success",
        "select_dir": "Select Directory",
        "watermark": "Produced by Wu Guanyu, Quality Guaranteed"
    }
}

class BatchRenameApp:
    def __init__(self, root):
        self.root = root
        self.lang = "zh"
        
        # 极客风格配色配置
        self.BG = "#0D1117"
        self.FG = "#39D353"
        self.ENTRY_BG = "#161B22"
        self.BTN_BG = "#21262D"
        self.BTN_ACT = "#30363D"
        self.FONT = ("Consolas", 10, "bold")
        self.TITLE_FONT = ("Consolas", 12, "bold")
        
        self.root.geometry("750x650")
        self.root.configure(bg=self.BG)
        
        self.setup_style()
        self.create_widgets()
        self.update_texts()
        
        # 迭代新增：设置程序内部运行时的图标
        self.set_app_icon()

    def set_app_icon(self):
        try:
            # 兼容判断：区分是在执行 .py 源码还是运行已打包后的 .exe 文件
            if getattr(sys, 'frozen', False):
                # 运行已打包的 .exe 文件
                base_path = os.path.dirname(sys.executable)
                # PyInstaller的临时运行目录（当资源通过 --add-data 打包在exe内部时）
                meipass_path = getattr(sys, '_MEIPASS', base_path)
            else:
                # 运行源码 .py 文件
                base_path = os.path.dirname(os.path.abspath(__file__))
                meipass_path = base_path

            # 探测路径 1：源码运行时，或是按照绝对路径放置在了 /dist/Resources/icon.jpg
            path1 = os.path.join(base_path, 'dist', 'Resources', 'icon.jpg')
            # 探测路径 2：打包后，作为外部文件夹和 exe 在同一级目录
            path2 = os.path.join(base_path, 'Resources', 'icon.jpg')
            # 探测路径 3：打包进单文件exe内部 (_MEIPASS/dist/Resources/icon.jpg)
            path3 = os.path.join(meipass_path, 'dist', 'Resources', 'icon.jpg')
            # 探测路径 4：打包进单文件exe内部 (_MEIPASS/Resources/icon.jpg)
            path4 = os.path.join(meipass_path, 'Resources', 'icon.jpg')

            icon_path = None
            for p in [path1, path2, path3, path4]:
                if os.path.exists(p):
                    icon_path = p
                    break

            # 若找到了图标且环境具备 Pillow 库，则应用图标及任务栏略缩图
            if icon_path and Image is not None and ImageTk is not None:
                # 【关键修复1】：强制通知 Windows 将应用程序的略缩图和状态栏图标与默认 Python/Tk 进程分离开来
                if os.name == 'nt':
                    import ctypes
                    try:
                        # 随便指定一个固定的AppID字符串，可确保任务栏不调用默认图标
                        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("batch.rename.tool.v1")
                    except Exception:
                        pass

                img = Image.open(icon_path)
                # 【关键修复2】：必须将 PhotoImage 对象持久化（绑定到 self 属性上），
                # 否则作为局部变量在函数结束后就会被 Python 垃圾回收，导致界面彻底不显示图标！
                self.app_icon_photo = ImageTk.PhotoImage(img)
                # True 表示将此图标应用至所有子窗口并向系统注册为任务栏图标
                self.root.iconphoto(True, self.app_icon_photo)
        except Exception:
            # 发生任何加载异常则静默跳过，绝不干扰正常原有功能
            pass

    def setup_style(self):
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("TCombobox", 
                        fieldbackground=self.ENTRY_BG, 
                        background=self.BTN_BG, 
                        foreground=self.FG, 
                        arrowcolor=self.FG,
                        darkcolor=self.BG,
                        lightcolor=self.BG,
                        selectbackground=self.FG,
                        selectforeground=self.BG)
        style.map("TCombobox", 
                  fieldbackground=[("readonly", self.ENTRY_BG)], 
                  foreground=[("readonly", self.FG)])

    def create_label(self, parent):
        return tk.Label(parent, bg=self.BG, fg=self.FG, font=self.FONT)

    def create_button(self, parent, command=None):
        return tk.Button(parent, bg=self.BTN_BG, fg=self.FG, activebackground=self.BTN_ACT, 
                         activeforeground=self.FG, font=self.FONT, relief=tk.SOLID, bd=1, command=command)

    def create_entry(self, parent, width=20, textvariable=None):
        return tk.Entry(parent, bg=self.ENTRY_BG, fg=self.FG, insertbackground=self.FG, 
                        font=self.FONT, relief=tk.SOLID, bd=1, width=width, textvariable=textvariable)

    def create_frame(self, parent):
        return tk.Frame(parent, bg=self.BG)

    def create_checkbutton(self, parent, variable=None, command=None):
        return tk.Checkbutton(parent, bg=self.BG, fg=self.FG, selectcolor=self.ENTRY_BG, 
                              activebackground=self.BG, activeforeground=self.FG, font=self.FONT,
                              variable=variable, command=command)

    def create_widgets(self):
        frame_top = self.create_frame(self.root)
        frame_top.pack(fill=tk.X, padx=20, pady=5)
        self.btn_lang = self.create_button(frame_top, command=self.toggle_lang)
        self.btn_lang.pack(side=tk.RIGHT)

        frame_input = self.create_frame(self.root)
        frame_input.pack(fill=tk.X, padx=20, pady=5)
        self.lbl_input_path = self.create_label(frame_input)
        self.lbl_input_path.pack(side=tk.LEFT)
        self.input_path_var = tk.StringVar()
        self.create_entry(frame_input, width=50, textvariable=self.input_path_var).pack(side=tk.LEFT, padx=5)
        self.btn_select_input = self.create_button(frame_input, command=self.select_input_path)
        self.btn_select_input.pack(side=tk.LEFT)

        frame_output = self.create_frame(self.root)
        frame_output.pack(fill=tk.X, padx=20, pady=5)
        self.same_path_var = tk.BooleanVar(value=True)
        self.chk_same_path = self.create_checkbutton(frame_output, variable=self.same_path_var, command=self.toggle_output_path)
        self.chk_same_path.pack(side=tk.LEFT)
        self.output_path_frame = self.create_frame(frame_output)
        self.lbl_output_path = self.create_label(self.output_path_frame)
        self.lbl_output_path.pack(side=tk.LEFT, padx=(10,0))
        self.output_path_var = tk.StringVar()
        self.create_entry(self.output_path_frame, width=30, textvariable=self.output_path_var).pack(side=tk.LEFT, padx=5)
        self.btn_select_output = self.create_button(self.output_path_frame, command=self.select_output_path)
        self.btn_select_output.pack(side=tk.LEFT)
        self.toggle_output_path()

        self.lf_settings = tk.LabelFrame(self.root, bg=self.BG, fg=self.FG, font=self.TITLE_FONT, bd=1, relief=tk.SOLID)
        self.lf_settings.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        frame_prefix = self.create_frame(self.lf_settings)
        frame_prefix.pack(fill=tk.X, padx=10, pady=10)
        self.use_prefix_var = tk.BooleanVar()
        self.chk_prefix = self.create_checkbutton(frame_prefix, variable=self.use_prefix_var, command=self.toggle_prefix_ui)
        self.chk_prefix.pack(side=tk.LEFT)
        self.prefix_ui_frame = self.create_frame(frame_prefix)
        self.prefix_combo = ttk.Combobox(self.prefix_ui_frame, state="readonly", width=12)
        self.prefix_combo.pack(side=tk.LEFT, padx=5)
        self.prefix_combo.bind("<<ComboboxSelected>>", self.toggle_prefix_mode)
        self.prefix_entry1 = self.create_entry(self.prefix_ui_frame, width=15)
        self.prefix_entry1.pack(side=tk.LEFT, padx=5)
        self.lbl_prefix_to = self.create_label(self.prefix_ui_frame)
        self.prefix_entry2 = self.create_entry(self.prefix_ui_frame, width=15)
        
        frame_suffix = self.create_frame(self.lf_settings)
        frame_suffix.pack(fill=tk.X, padx=10, pady=10)
        self.use_suffix_var = tk.BooleanVar()
        self.chk_suffix = self.create_checkbutton(frame_suffix, variable=self.use_suffix_var, command=self.toggle_suffix_ui)
        self.chk_suffix.pack(side=tk.LEFT)
        self.suffix_ui_frame = self.create_frame(frame_suffix)
        self.suffix_combo = ttk.Combobox(self.suffix_ui_frame, state="readonly", width=12)
        self.suffix_combo.pack(side=tk.LEFT, padx=5)
        self.suffix_combo.bind("<<ComboboxSelected>>", self.toggle_suffix_mode)
        self.suffix_entry1 = self.create_entry(self.suffix_ui_frame, width=15)
        self.suffix_entry1.pack(side=tk.LEFT, padx=5)
        self.lbl_suffix_to = self.create_label(self.suffix_ui_frame)
        self.suffix_entry2 = self.create_entry(self.suffix_ui_frame, width=15)

        frame_seq_cust = self.create_frame(self.lf_settings)
        frame_seq_cust.pack(fill=tk.X, padx=10, pady=10)
        self.use_seq_var = tk.BooleanVar()
        self.chk_seq = self.create_checkbutton(frame_seq_cust, variable=self.use_seq_var)
        self.chk_seq.pack(side=tk.LEFT)
        self.use_custom_var = tk.BooleanVar()
        self.chk_custom = self.create_checkbutton(frame_seq_cust, variable=self.use_custom_var, command=self.toggle_custom_ui)
        self.chk_custom.pack(side=tk.LEFT, padx=(20,0))
        self.custom_ui_frame = self.create_frame(frame_seq_cust)
        self.lbl_custom = self.create_label(self.custom_ui_frame)
        self.lbl_custom.pack(side=tk.LEFT)
        self.custom_entry = self.create_entry(self.custom_ui_frame, width=20)
        self.custom_entry.pack(side=tk.LEFT, padx=5)

        frame_export = self.create_frame(self.root)
        frame_export.pack(fill=tk.X, padx=20, pady=10)
        self.btn_export = self.create_button(frame_export, command=self.export_filenames)
        self.btn_export.pack(side=tk.LEFT)
        self.lbl_export_fmt = self.create_label(frame_export)
        self.lbl_export_fmt.pack(side=tk.LEFT, padx=(15,5))
        self.export_format_var = tk.StringVar(value=".xlsx")
        ttk.Combobox(frame_export, textvariable=self.export_format_var, values=[".xlsx", ".txt", ".csv"], state="readonly", width=8).pack(side=tk.LEFT)
        self.include_ext_var = tk.BooleanVar(value=True)
        self.chk_include_ext = self.create_checkbutton(frame_export, variable=self.include_ext_var)
        self.chk_include_ext.pack(side=tk.LEFT, padx=(15,0))

        self.btn_run = self.create_button(self.root, command=self.run_rename)
        self.btn_run.pack(pady=20, ipadx=10, ipady=5)

        # 增加本地化水印组件
        self.lbl_watermark = tk.Label(self.root, bg=self.BG, fg=self.FG, font=("Consolas", 9, "italic"))
        self.lbl_watermark.pack(side=tk.BOTTOM, pady=10)

    def toggle_lang(self):
        self.lang = "en" if self.lang == "zh" else "zh"
        self.update_texts()

    def update_texts(self):
        t = LANG[self.lang]
        self.root.title(t["title"])
        self.btn_lang.config(text=t["lang_btn"])
        
        self.lbl_input_path.config(text=t["input_path"])
        self.btn_select_input.config(text=t["select_path"])
        self.chk_same_path.config(text=t["same_path"])
        self.lbl_output_path.config(text=t["output_path"])
        self.btn_select_output.config(text=t["select_path"])
        self.lf_settings.config(text=t["custom_settings"])
        
        self.chk_prefix.config(text=t["add_mod_prefix"])
        self.lbl_prefix_to.config(text=t["replace_with"])
        self.chk_suffix.config(text=t["add_mod_suffix"])
        self.lbl_suffix_to.config(text=t["replace_with"])
        self.chk_seq.config(text=t["seq_naming"])
        self.chk_custom.config(text=t["custom_naming"])
        self.lbl_custom.config(text=t["new_name"])
        
        self.btn_export.config(text=t["export_btn"])
        self.lbl_export_fmt.config(text=t["export_format"])
        self.chk_include_ext.config(text=t["include_ext"])
        self.btn_run.config(text=t["run_btn"])
        
        p_idx = self.prefix_combo.current() if self.prefix_combo.current() != -1 else 0
        self.prefix_combo.config(values=t["prefix_modes"])
        self.prefix_combo.current(p_idx)
        
        s_idx = self.suffix_combo.current() if self.suffix_combo.current() != -1 else 0
        self.suffix_combo.config(values=t["suffix_modes"])
        self.suffix_combo.current(s_idx)

        # 更新水印文本
        self.lbl_watermark.config(text=t["watermark"])

    def select_input_path(self):
        path = filedialog.askdirectory(title=LANG[self.lang]["select_dir"])
        if path: self.input_path_var.set(path)

    def select_output_path(self):
        path = filedialog.askdirectory(title=LANG[self.lang]["select_dir"])
        if path: self.output_path_var.set(path)

    def toggle_output_path(self):
        if self.same_path_var.get():
            self.output_path_frame.pack_forget()
        else:
            self.output_path_frame.pack(side=tk.LEFT, padx=10)

    def toggle_prefix_ui(self):
        if self.use_prefix_var.get():
            self.prefix_ui_frame.pack(side=tk.LEFT, padx=10)
            self.toggle_prefix_mode()
        else:
            self.prefix_ui_frame.pack_forget()

    def toggle_prefix_mode(self, event=None):
        if self.prefix_combo.current() == 1:
            self.lbl_prefix_to.pack(side=tk.LEFT, padx=5)
            self.prefix_entry2.pack(side=tk.LEFT, padx=5)
        else:
            self.lbl_prefix_to.pack_forget()
            self.prefix_entry2.pack_forget()

    def toggle_suffix_ui(self):
        if self.use_suffix_var.get():
            self.suffix_ui_frame.pack(side=tk.LEFT, padx=10)
            self.toggle_suffix_mode()
        else:
            self.suffix_ui_frame.pack_forget()

    def toggle_suffix_mode(self, event=None):
        if self.suffix_combo.current() == 1:
            self.lbl_suffix_to.pack(side=tk.LEFT, padx=5)
            self.suffix_entry2.pack(side=tk.LEFT, padx=5)
        else:
            self.lbl_suffix_to.pack_forget()
            self.suffix_entry2.pack_forget()

    def toggle_custom_ui(self):
        if self.use_custom_var.get():
            self.custom_ui_frame.pack(side=tk.LEFT, padx=10)
            self.use_seq_var.set(True)
            self.chk_seq.config(state=tk.DISABLED)
        else:
            self.custom_ui_frame.pack_forget()
            self.chk_seq.config(state=tk.NORMAL)

    def get_target_files(self, folder):
        t = LANG[self.lang]
        try:
            # 优化点：使用 os.scandir 替代 os.listdir + isfile，极大提升大目录遍历效率
            with os.scandir(folder) as it:
                return [entry.name for entry in it if entry.is_file()]
        except Exception as e:
            messagebox.showerror(t["error"], t["msg_err"].format(e))
            return []

    def export_filenames(self):
        t = LANG[self.lang]
        input_dir = self.input_path_var.get()
        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showwarning(t["warning"], t["msg_warn_input"])
            return

        files = self.get_target_files(input_dir)
        if not files:
            messagebox.showinfo(t["info"], t["msg_no_files"])
            return

        fmt = self.export_format_var.get()
        export_path = filedialog.asksaveasfilename(
            initialdir=input_dir,
            initialfile=t["file_list_name"],
            defaultextension=fmt,
            filetypes=[(f"{fmt} Files", f"*{fmt}"), ("All Files", "*.*")],
            title=t["export_btn"]
        )

        if not export_path:
            return

        include_ext = self.include_ext_var.get()
        # 优化点：使用生成器表达式代替列表生成，减少瞬时内存占用
        out_files = (f if include_ext else os.path.splitext(f)[0] for f in files)
        header = t["filename_header"]

        try:
            if fmt == ".txt":
                with open(export_path, 'w', encoding='utf-8') as f:
                    for file in out_files:
                        f.write(file + "\n")
            elif fmt == ".csv":
                with open(export_path, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([header])
                    for file in out_files:
                        writer.writerow([file])
            elif fmt == ".xlsx":
                if openpyxl is None:
                    messagebox.showerror(t["error"], t["msg_no_openpyxl"])
                    return
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append([header])
                for file in out_files:
                    ws.append([file])
                wb.save(export_path)

            messagebox.showinfo(t["success"], t["msg_export_succ"].format(export_path))
        except Exception as e:
            messagebox.showerror(t["error"], t["msg_err"].format(e))

    def run_rename(self):
        t = LANG[self.lang]
        input_dir = self.input_path_var.get()
        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showwarning(t["warning"], t["msg_warn_input"])
            return

        same_path = self.same_path_var.get()
        output_dir = input_dir if same_path else self.output_path_var.get()
        
        if not same_path:
            if not output_dir:
                messagebox.showwarning(t["warning"], t["msg_warn_output"])
                return
            # 优化点：添加 exist_ok=True 避免由于并发产生的目录创建竞争崩溃
            os.makedirs(output_dir, exist_ok=True)

        files = self.get_target_files(input_dir)
        if not files:
            messagebox.showinfo(t["info"], t["msg_no_files"])
            return

        # ================= 性能优化核心区 =================
        # 优化点：将 Tkinter 控件的取值操作从 for 循环中提取出来。
        # 之前每次循环都要调用底层 Tk 引擎数十次，现在 O(1) 取值。
        
        use_custom = self.use_custom_var.get()
        custom_name = self.custom_entry.get()
        
        use_prefix = self.use_prefix_var.get()
        prefix_mode = self.prefix_combo.current()
        prefix_1 = self.prefix_entry1.get()
        prefix_2 = self.prefix_entry2.get()
        
        use_suffix = self.use_suffix_var.get()
        suffix_mode = self.suffix_combo.current()
        suffix_1 = self.suffix_entry1.get()
        suffix_2 = self.suffix_entry2.get()
        
        use_seq = self.use_seq_var.get()
        # ===============================================

        success_count = 0
        
        for index, filename in enumerate(files, start=1):
            name, ext = os.path.splitext(filename)
            
            # 使用了三元运算与 f-string 代替传统 + 号拼接
            new_name = custom_name if use_custom else name

            if use_prefix:
                if prefix_mode == 0:
                    new_name = f"{prefix_1}{new_name}"
                elif prefix_1: 
                    new_name = new_name.replace(prefix_1, prefix_2)

            if use_suffix:
                if suffix_mode == 0:
                    new_name = f"{new_name}{suffix_1}"
                elif suffix_1: 
                    new_name = new_name.replace(suffix_1, suffix_2)

            if use_seq:
                new_name = f"{new_name}_{index}"

            final_filename = f"{new_name}{ext}"
            src_path = os.path.join(input_dir, filename)
            dst_path = os.path.join(output_dir, final_filename)

            # 漏洞修复：如果用户什么设置都没勾选就点击运行，原逻辑会让自身对自身 rename 导致报错或者文件损毁
            if src_path == dst_path:
                continue

            try:
                if same_path:
                    os.rename(src_path, dst_path)
                else:
                    shutil.copy2(src_path, dst_path)
                success_count += 1
            except Exception as e:
                # 遇到错误时不中断整体流程，仅输出记录（可依据需求转为日志系统）
                print(f"Error handling file {filename}: {e}")

        messagebox.showinfo(t["success"], t["msg_done"].format(success_count))


if __name__ == "__main__":
    root = tk.Tk()
    app = BatchRenameApp(root)
    root.mainloop()