import os
import subprocess
import sys
import shutil
import json
from datetime import datetime
from pydub import AudioSegment
from pydub.exceptions import CouldntDecodeError
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from tkinter import scrolledtext
from natsort import natsorted
import warnings
warnings.filterwarnings("ignore", category=RuntimeWarning, module="pydub")


def resource_path(relative_path):
    """获取资源的绝对路径（兼容开发环境和打包后）"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class DraggableListbox(tk.Listbox):
    """支持拖拽排序和右键菜单的Listbox组件"""

    def __init__(self, master, app, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app
        self.bind('<Button-1>', self.set_current)
        self.bind('<B1-Motion>', self.shift_selection)
        self.bind('<Button-3>', self.show_context_menu)
        self.curIndex = None

        # 创建右键菜单
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="删除", command=self.delete_selected)
        self.context_menu.add_command(label="上移", command=self.move_up)
        self.context_menu.add_command(label="下移", command=self.move_down)

    def set_current(self, event):
        self.curIndex = self.nearest(event.y)

    def shift_selection(self, event):
        i = self.nearest(event.y)
        if i < self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i + 1, x)
            self.curIndex = i
        elif i > self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i - 1, x)
            self.curIndex = i

    def show_context_menu(self, event):
        """显示右键菜单"""
        try:
            index = self.nearest(event.y)
            self.selection_clear(0, tk.END)
            self.selection_set(index)
            self.context_menu.post(event.x_root, event.y_root)
        except:
            pass

    def delete_selected(self):
        """删除选中项"""
        selection = self.curselection()
        if selection:
            self.delete(selection[0])
            self.app.log("已从合并列表中移除一个文件")

    def move_up(self):
        """上移选中项"""
        selection = self.curselection()
        if selection and selection[0] > 0:
            index = selection[0]
            item = self.get(index)
            self.delete(index)
            self.insert(index - 1, item)
            self.selection_set(index - 1)

    def move_down(self):
        """下移选中项"""
        selection = self.curselection()
        if selection and selection[0] < self.size() - 1:
            index = selection[0]
            item = self.get(index)
            self.delete(index)
            self.insert(index + 1, item)
            self.selection_set(index + 1)


class AudioMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("音频文件合并工具")
        self.root.geometry("800x700+100+20")
        self.root.resizable(False, False)

        # 初始化变量
        self.processing = False
        self.audio_files = []

        # 配置文件路径
        self.config_file = self.get_config_path()

        # 设置主题
        self.style = ttk.Style()
        # 为Windows系统设置启动信息以隐藏命令行窗口
        if sys.platform.startswith('win'):
            self.startupinfo = subprocess.STARTUPINFO()
            self.startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        else:
            self.startupinfo = None

        # 创建主界面
        self.create_widgets()

        # 初始化FFmpeg路径（必须在界面创建后调用）
        self.setup_ffmpeg()

        # 加载配置
        self.load_config()

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def get_config_path(self):
        """获取配置文件路径"""
        if sys.platform.startswith('win'):
            config_dir = os.path.join(os.environ.get('APPDATA', '.'), 'AudioMerger')
        else:
            config_dir = os.path.join(os.path.expanduser('~'), '.AudioMerger')

        # 创建配置目录（如果不存在）
        os.makedirs(config_dir, exist_ok=True)
        return os.path.join(config_dir, 'config.json')

    def save_config(self):
        """保存配置到文件"""
        try:
            # 获取窗口位置
            x = self.root.winfo_x()
            y = self.root.winfo_y()

            # 获取当前主题
            current_theme = self.theme_var.get()

            # 配置数据
            config_data = {
                'window_position': (x, y),
                'theme': current_theme
            }

            # 保存到文件
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)

            self.log("配置已保存")
        except Exception as e:
            self.log(f"保存配置失败: {str(e)}")

    def load_config(self):
        """从文件加载配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)

                # 恢复窗口位置
                if 'window_position' in config_data:
                    x, y = config_data['window_position']
                    self.root.geometry(f"+{x}+{y}")
                    self.log(f"已恢复窗口位置")

                # 恢复主题
                if 'theme' in config_data and config_data['theme'] in self.style.theme_names():
                    self.theme_var.set(config_data['theme'])
                    self.style.theme_use(config_data['theme'])
                    self.log(f"已恢复 {config_data['theme']} 主题")

        except Exception as e:
            self.log(f"加载配置失败: {str(e)}")

    def on_close(self):
        """窗口关闭时的处理"""
        self.save_config()
        self.root.destroy()

    def configure_styles(self):
        """配置界面样式"""
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('微软雅黑', 9))
        self.style.configure('TButton', font=('微软雅黑', 9))
        self.style.configure('TEntry', font=('Consolas', 9))
        self.style.configure('TCombobox', font=('Consolas', 9))
        self.style.configure('TLabelFrame', background='#f0f0f0', font=('微软雅黑', 9, 'bold'))
        self.style.configure('Vertical.TScrollbar', background='#e0e0e0')

    def setup_ffmpeg(self):
        """配置FFmpeg路径"""
        try:
            ffmpeg_path = resource_path("ffmpeg.exe")
            if os.path.exists(ffmpeg_path):
                AudioSegment.ffmpeg = ffmpeg_path
                AudioSegment.converter = ffmpeg_path
                self.log("FFmpeg已从程序目录加载")
            else:
                self.log("警告: 未找到内置FFmpeg,请确保系统已安装FFmpeg并配置环境变量")
        except Exception as e:
            self.log(f"FFmpeg配置错误: {str(e)}")

    def change_theme(self, event=None):
        """切换主题"""
        selected_theme = self.theme_var.get()
        self.style.theme_use(selected_theme)
        self.log(f"已切换至 {selected_theme} 主题")
        # 重新配置样式以适应新主题
        self.configure_styles()

    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件夹选择部分
        folder_frame = ttk.LabelFrame(main_frame, text=" 文件夹选择 ", padding=10)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        self.folder_path = tk.StringVar()
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_path, font=('微软雅黑', 9))
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = ttk.Button(folder_frame, text="浏览...", command=self.browse_folder)
        browse_btn.pack(side=tk.RIGHT)

        # 排序和主题选项部分
        options_frame = ttk.LabelFrame(main_frame, text=" 选项设置 ", padding=10)
        options_frame.pack(fill=tk.X, pady=(0, 10))

        # 排序方式选择
        sort_frame = ttk.Frame(options_frame)
        sort_frame.pack(side=tk.LEFT, padx=10)
        ttk.Label(sort_frame, text="排序方式:").pack(side=tk.LEFT)
        self.sort_method = tk.StringVar(value="ctime")
        ttk.Radiobutton(sort_frame, text="创建时间", variable=self.sort_method,
                        value="ctime", command=self.update_file_list).pack(side=tk.LEFT)
        ttk.Radiobutton(sort_frame, text="文件名", variable=self.sort_method,
                        value="name", command=self.update_file_list).pack(side=tk.LEFT)

        # 主题选择
        theme_frame = ttk.Frame(options_frame)
        theme_frame.pack(side=tk.LEFT, padx=10)
        ttk.Label(theme_frame, text="界面主题:").pack(side=tk.LEFT)

        # 获取可用主题
        available_themes = self.style.theme_names()
        self.theme_var = tk.StringVar()
        theme_combo = ttk.Combobox(theme_frame, textvariable=self.theme_var,
                                   values=available_themes, width=10)
        theme_combo.pack(side=tk.LEFT)

        # 设置默认主题
        default_theme = "clam" if "clam" in available_themes else available_themes[0]
        self.theme_var.set(default_theme)
        self.style.theme_use(default_theme)

        # 绑定主题切换事件
        theme_combo.bind("<<ComboboxSelected>>", self.change_theme)

        # 文件列表显示
        list_frame = ttk.LabelFrame(main_frame, text=" 音频文件列表 (可拖拽调整顺序) ", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        list_frame.pack_propagate(False)
        list_frame.config(height=120)
        self.setup_icon();

        self.file_listbox = DraggableListbox(
            list_frame,
            self,
            selectmode=tk.SINGLE,
            activestyle='none',
            font=('Consolas', 9),
            background='white',
            selectbackground='#4a98db',
            selectforeground='white'
        )
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.pack(fill=tk.BOTH, expand=True)

        # 输出选项部分
        output_frame = ttk.LabelFrame(main_frame, text=" 输出选项 ", padding=10)
        output_frame.pack(fill=tk.X, pady=(0, 10))

        # 输出格式
        ttk.Label(output_frame, text="输出格式:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.output_format = tk.StringVar(value="mp3")
        format_combo = ttk.Combobox(output_frame, textvariable=self.output_format,
                                    values=('mp3', 'wav', 'ogg', 'flac'), width=8)
        format_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))

        # 输出文件名
        ttk.Label(output_frame, text="输出文件名:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.output_name = tk.StringVar(value="merged_audio")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_name, width=25)
        output_entry.grid(row=0, column=3, sticky=tk.W)

        # 进度条
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(10, 5))

        # 日志框
        log_frame = ttk.LabelFrame(main_frame, text=" 处理日志 ", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        log_frame.pack_propagate(False)
        log_frame.config(height=90)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            font=('Consolas', 9),
            padx=5,
            pady=5
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))

        self.process_btn = ttk.Button(button_frame, text="开始合并", command=self.start_processing)
        self.process_btn.pack(side=tk.RIGHT, padx=(10, 0))

        self.cancel_btn = ttk.Button(button_frame, text="取消", command=self.cancel_processing, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.RIGHT)

    def setup_icon(self):
        """设置应用图标"""
        try:
            icon_path = resource_path("catall3.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            self.log(f"图标加载错误: {str(e)}")

    def browse_folder(self):
        """选择文件夹并更新文件列表"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.log(f"已选择文件夹: {folder_selected}")
            self.update_file_list()

    def update_file_list(self):
        """更新文件列表显示"""
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            return

        self.file_listbox.delete(0, tk.END)
        self.audio_files = []
        supported_formats = ('.mp3', '.wav', '.ogg', '.flac', '.m4a', '.aac')

        # 收集文件信息
        for file in os.listdir(folder):
            if file.lower().endswith(supported_formats):
                file_path = os.path.join(folder, file)
                ctime = os.path.getctime(file_path)
                mtime = os.path.getmtime(file_path)
                self.audio_files.append({
                    'path': file_path,
                    'name': file,
                    'ctime': ctime,
                    'mtime': mtime
                })

        if not self.audio_files:
            self.log("没有找到支持的音频文件")
            return

        # 根据选择的排序方式排序
        sort_method = self.sort_method.get()
        if sort_method == "ctime":
            self.audio_files.sort(key=lambda x: x['ctime'])
            self.log("已按创建时间排序")
        else:
            self.audio_files = natsorted(self.audio_files, key=lambda x: x['name'])
            self.log("已按文件名排序")

        # 显示文件列表
        for file_info in self.audio_files:
            timestamp = datetime.fromtimestamp(file_info['ctime']).strftime('%Y-%m-%d %H:%M:%S')
            self.file_listbox.insert(tk.END, f"{file_info['name']} (创建时间: {timestamp})")

        self.log(f"找到 {len(self.audio_files)} 个音频文件")

    def log(self, message):
        """记录日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def start_processing(self):
        """开始处理音频合并"""
        if not self.audio_files:
            messagebox.showerror("错误", "没有可合并的音频文件")
            return

        # 禁用按钮，启用取消按钮
        self.process_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)

        # 获取用户调整后的顺序
        ordered_files = []
        for i in range(self.file_listbox.size()):
            display_text = self.file_listbox.get(i)
            # 从显示文本中提取文件名
            file_name = display_text.split(' (创建时间:')[0]
            # 找到对应的文件信息
            for file_info in self.audio_files:
                if file_info['name'] == file_name:
                    ordered_files.append(file_info['path'])
                    break

        # 在后台线程中运行处理过程
        self.processing = True
        threading.Thread(
            target=self.process_audio,
            args=(ordered_files,),
            daemon=True
        ).start()

    def cancel_processing(self):
        """取消处理"""
        self.processing = False
        self.log("用户请求取消处理...")

    def process_audio(self, file_paths):
        """实际处理音频合并"""
        try:
            folder = self.folder_path.get()
            output_format = self.output_format.get()
            output_name = self.output_name.get() + "." + output_format
            output_path = os.path.join(folder, output_name)

            # 为Windows系统设置启动信息以隐藏命令行窗口
            if sys.platform.startswith('win'):
                import subprocess
                from functools import wraps

                # 创建启动信息以隐藏窗口
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

                # 保存原始的subprocess.Popen
                original_popen = subprocess.Popen

                # 创建包装函数
                @wraps(original_popen)
                def subprocess_popen(cmd, **kwargs):
                    # 确保startupinfo参数存在
                    if 'startupinfo' not in kwargs:
                        kwargs['startupinfo'] = startupinfo
                    # 调用原始的Popen
                    return original_popen(cmd, **kwargs)

                # 临时替换subprocess.Popen
                subprocess.Popen = subprocess_popen

            combined = None
            total_files = len(file_paths)

            for i, file_path in enumerate(file_paths):
                if not self.processing:
                    self.log("处理已取消")
                    return

                try:
                    self.log(f"正在处理: {os.path.basename(file_path)} ({i + 1}/{total_files})")
                    audio = AudioSegment.from_file(file_path)

                    if combined is None:
                        combined = audio
                    else:
                        combined += audio

                    self.progress['value'] = (i + 1) / total_files * 100
                except CouldntDecodeError:
                    self.log(f"警告: 无法解码文件 {os.path.basename(file_path)}，已跳过")
                except Exception as e:
                    self.log(f"处理文件时出错 {os.path.basename(file_path)}: {str(e)}")

            if combined is None:
                self.log("错误: 没有有效的音频文件可合并")
                messagebox.showerror("错误", "没有有效的音频文件可合并")
                return

            # 导出合并后的文件
            self.log("正在导出合并后的音频文件...")
            combined.export(output_path, format=output_format)
            self.log(f"已成功导出: {output_name}")

            messagebox.showinfo("完成", f"音频文件已成功合并为: {output_name}")

        except Exception as e:
            self.log(f"处理过程中出错: {str(e)}")
            messagebox.showerror("错误", f"处理过程中出错: {str(e)}")
        finally:
            # 重置UI状态
            self.process_btn.config(state=tk.NORMAL)
            self.cancel_btn.config(state=tk.DISABLED)
            self.processing = False
            self.progress['value'] = 0
    def on_close(self):
        """窗口关闭时保存配置并退出"""
        self.save_config()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = AudioMergerApp(root)
    root.mainloop()
