# 导入必要的模块
import locale
import os
import sys
import subprocess
import shutil
import threading
import queue
import tkinter as tk
import random
import string
import re
from tkinter import scrolledtext

class AutoUnzipApp:
    def __init__(self):
        # 初始化日志队列，用于线程间安全传递日志信息
        self.log_queue = queue.Queue()
        # 初始化任务队列，用于存储待处理的文件路径
        self.task_queue = queue.Queue()
        # 创建线程锁，用于保护共享资源的并发访问
        self.lock = threading.Lock()
        # 创建主窗口
        self.window = tk.Tk()
        # 设置窗口标题
        self.window.title('auto_unzip')
        # 设置窗口大小为800x600像素
        self.window.geometry('800x600')
        # 创建滚动文本区域用于显示日志信息，设置自动换行
        self.log_area = scrolledtext.ScrolledText(self.window, wrap=tk.WORD)
        # 打包滚动文本区域使其填充整个窗口
        self.log_area.pack(expand=True, fill='both')
        # 创建状态框架
        self.status_frame = tk.Frame(self.window)
        # 打包状态框架使其横向填充
        self.status_frame.pack(fill=tk.X)
        # 创建状态标签，初始显示解压次数为0
        self.status_label = tk.Label(self.status_frame, text='解压次数: 0')
        # 将状态标签放在状态框架的左侧
        self.status_label.pack(side=tk.LEFT)
        self.log('本工具已开源在https://github.com/liaoran2002/auto_unzip')
        # 活动任务数计数器
        self.active_tasks = 0
        # 解压次数计数器
        self.process_count = 0
        # 处理单个文件的最大尝试次数
        self.max_attempts = 10
        # 最大工作线程数
        self.max_workers = 4
        # 临时目录前缀
        self.temp_dir_prefix = 'tmp_'  # 从 'tmo_' 改为 'tmp_'，更符合标准
        # 支持的压缩文件扩展名集合
        self.compressed_exts = {'7z', 'rar', 'zip'}
        # 分卷文件正则表达式，用于匹配分卷压缩文件
        self.split_file_pattern = re.compile(r'(\.part\d+\.rar|\.r\d{2,}|\.7z\.\d{3,}|\.\d{3,})$', re.IGNORECASE)
        # 密码文件路径
        self.passwords_file_path = os.path.join(self.get_exe_dir(),'passwords.txt')
        # 密码列表
        self.passwords = []

        # 加载密码文件
        if os.path.exists(self.passwords_file_path):
            # 定义尝试的编码列表，优先使用UTF-8，其次是系统默认编码和GBK
            encodings = ['utf-8', locale.getpreferredencoding(), 'gbk']
            for encoding in encodings:
                try:
                    # 尝试使用当前编码打开并读取密码文件
                    with open(self.passwords_file_path, 'r', encoding=encoding) as f:
                        # 读取所有非空行作为密码列表
                        self.passwords = [line.strip() for line in f if line.strip()]
                    # 成功读取后记录日志并退出循环
                    self.log(f'成功使用 {encoding} 编码加载 {len(self.passwords)} 个密码: {self.passwords}')
                    break
                except UnicodeDecodeError:
                    # 如果当前编码解码失败，尝试下一个编码
                    continue
                except Exception as e:
                    # 如果发生其他错误，记录日志并退出循环
                    self.log(f'使用 {encoding} 编码加载密码文件时出错: {e}')
                    break
            else:
                # 如果所有编码都尝试失败，清空密码列表并记录日志
                self.passwords = []
                self.log('所有编码尝试失败，无法加载密码文件')
        else:
            # 如果密码文件不存在，创建空的密码文件
            with open(self.passwords_file_path, 'w', encoding='utf-8') as f:
                f.write('')
            # 创建示例密码文件
            with open(os.path.join(self.get_exe_dir(), 'passwords_example.txt'), 'w', encoding='utf-8') as f:
                f.write('666\n888\n嘿嘿嘿\n')
            # 记录日志提示用户如何配置密码文件
            self.log('未找到密码文件passwords.txt，已创建空文件')
            self.log('请参考passwords_example.txt中的示例在passwords.txt中每行添加一个密码')

        # 启动指定数量的工作线程
        for _ in range(self.max_workers):
            threading.Thread(target=self.process_worker, daemon=True).start()

        # 每隔100毫秒检查一次日志队列
        self.window.after(100, self.check_log_queue)
        
        # 如果有命令行参数，将其作为任务添加
        if len(sys.argv) > 1:
            for file_path in sys.argv[1:]:
                self.add_task(file_path)
        else:
            self.log("请将文件拖到文件图标上使用")
            self.window.after(1000, self.window.destroy)


    def get_7z_path(self):
        # 获取7z可执行文件的路径
        if getattr(sys, 'frozen', False):
            # 如果是打包后的可执行文件
            base_path = sys._MEIPASS
        else:
            # 如果是开发时的脚本
            base_path = os.path.dirname(os.path.abspath(__file__))  # 开发时的项目路径
        return os.path.join(base_path, '7z')

    def get_exe_dir(self):
        # 获取当前可执行文件或脚本的目录
        if getattr(sys, 'frozen', False):
            # 如果是打包后的可执行文件
            exe_path = sys.executable
        else:
            # 如果是开发时的脚本
            exe_path = os.path.abspath(__file__)
        return os.path.dirname(exe_path)

    def generate_temp_dir(self):
        # 生成随机临时目录名
        chars = string.ascii_lowercase + string.digits
        return self.temp_dir_prefix + ''.join(random.choices(chars, k=5))

    def check_log_queue(self):
        # 检查日志队列并更新日志显示
        if not self.log_queue.empty():
            msg = self.log_queue.get()
            self.log_area.insert(tk.END, msg + '\n')
            self.log_area.see(tk.END)  # 滚动到日志末尾
        # 100毫秒后再次检查
        self.window.after(100, self.check_log_queue)

    def log(self, message):
        # 将日志消息放入日志队列
        self.log_queue.put(message)

    def update_status(self):
        # 更新状态标签显示的解压次数
        self.status_label.config(text=f'解压次数: {self.process_count}')

    def add_task(self, file_path):
        # 添加文件路径到任务队列
        abs_path = os.path.abspath(file_path)  # 获取绝对路径
        if not os.path.exists(abs_path):
            self.log(f'文件不存在: {abs_path}')
            return
        with self.lock:  # 线程安全地访问任务队列
            self.task_queue.put(abs_path)
            self.log(f'已添加任务: {abs_path}')
            self.active_tasks += 1  # 活动任务数加1

    def process_worker(self):
        # 工作线程函数
        while True:
            current_path = self.task_queue.get()  # 从任务队列获取文件路径
            try:
                self.process_single_file(current_path)  # 处理单个文件
            finally:
                self.task_queue.task_done()  # 标记任务完成
                with self.lock:  # 线程安全地更新活动任务数
                    self.active_tasks -= 1  # 活动任务数减1
                    if self.active_tasks == 0:
                        # 如果所有任务都已完成，1秒后关闭窗口
                        self.window.after(1000, self.window.destroy)

    def is_compressed_file(self, filename):
        # 检查文件是否为压缩文件或分卷文件
        base_name = os.path.basename(filename)  # 获取文件名
        ext = base_name.split('.')[-1].lower()  # 获取文件扩展名并转为小写
        if ext in self.compressed_exts:
            # 如果扩展名在支持的压缩扩展名集合中
            return True
        # 否则检查是否为分卷文件
        return self.split_file_pattern.search(base_name) is not None

    def process_single_file(self, initial_path):
        # 处理单个文件
        file_queue = queue.Queue()  # 创建文件队列
        file_queue.put(initial_path)  # 将初始路径放入队列
        local_count = 0  # 本地解压次数计数器
        
        while not file_queue.empty() and local_count < self.max_attempts:
            current_path = file_queue.get()  # 获取当前处理路径
            if os.path.isdir(current_path):
                # 如果是目录
                self.log(f'找到目录: {os.path.basename(current_path)}')
                continue
            original_name = os.path.basename(current_path)  # 获取原始文件名
            if self.is_compressed_file(current_path):
                # 如果是压缩文件
                self.log(f'检测到压缩文件: {original_name}')
                extracted_files = self.extract_file(current_path)  # 解压文件
                # 处理解压结果
                success = self.handle_extraction_result(current_path, extracted_files, file_queue)
            else:
                # 如果是非压缩文件
                self.log(f'检测到非压缩文件: {original_name}，尝试重命名为压缩文件格式')
                # 尝试重命名为压缩文件格式并处理
                success = self.process_as_non_compressed(current_path, file_queue)
            if success:
                local_count += 1
            else:
                self.log(f'文件无法解压: {current_path}')
        self.log(f'文件 {initial_path} 处理完成：达到最大次数或找到目录')

    def process_as_non_compressed(self, current_path, file_queue):
        # 将非压缩文件视为压缩文件处理
        original_name = os.path.basename(current_path)  # 获取原始文件名
        for ext in self.compressed_exts:
            # 尝试为文件添加不同的压缩文件扩展名
            temp_file = f'{current_path}.{ext}'
            with self.lock:
                if not os.path.exists(current_path):
                    self.log(f'文件已被其他进程处理: {original_name}')
                    continue
                try:
                    os.rename(current_path, temp_file)  # 重命名文件
                except Exception as e:
                    self.log(f'重命名失败: {str(e)}')
                    continue
            self.log(f'尝试解压: {temp_file}')
            extracted_files = self.extract_file(temp_file)  # 尝试解压
            if extracted_files:
                # 如果解压成功，处理解压结果
                self.handle_extraction_result(temp_file, extracted_files, file_queue)
                return True
            else:
                # 如果解压失败，恢复原始文件名
                with self.lock:
                    if os.path.exists(temp_file):
                        os.rename(temp_file, current_path)
        return False

    def find_split_files(self, main_file):
        # 查找分卷文件
        base_name = os.path.basename(main_file)  # 获取文件名
        match = self.split_file_pattern.search(base_name)  # 检查是否为分卷文件
        if not match:
            return []
        
        base_pattern = self.split_file_pattern.sub('', base_name)  # 获取分卷文件的基础名称
        dir_path = os.path.dirname(main_file)  # 获取文件所在目录
        split_files = []
        
        for f in os.listdir(dir_path):
            # 遍历目录下的所有文件
            if f.startswith(base_pattern) and f != base_name:
                # 如果文件名以基础名称开头且不是主文件
                full_path = os.path.join(dir_path, f)  # 获取完整路径
                if self.split_file_pattern.search(f):
                    # 如果是分卷文件
                    split_files.append(full_path)
        split_files.sort()  # 排序分卷文件
        return split_files

    def handle_extraction_result(self, src_file, extracted_files, file_queue):
        # 处理解压结果
        if not extracted_files:
            # 如果没有提取到文件
            return False
        split_files = self.find_split_files(src_file)  # 查找分卷文件
        if split_files:
            # 如果找到分卷文件
            self.log(f'发现分卷文件组：共 {len(split_files)} 个')
        deleted_files = set()  # 已删除文件集合
        for f in split_files:
            with self.lock:
                try:
                    if os.path.exists(f) and f not in deleted_files:
                        os.remove(f)  # 删除分卷文件
                        deleted_files.add(f)
                        self.log(f'删除分卷文件: {os.path.basename(f)}')
                except Exception as e:
                    self.log(f'删除分卷失败: {str(e)}')
        with self.lock:
            try:
                if os.path.exists(src_file) and src_file not in deleted_files:
                    os.remove(src_file)  # 删除原压缩文件
                    deleted_files.add(src_file)
                    self.log(f'删除原文件: {os.path.basename(src_file)}')
            except Exception as e:
                self.log(f'删除原文件失败: {str(e)}')
        self.process_count += 1  # 解压次数加1
        self.window.event_generate('<<UpdateStatus>>')  # 触发状态更新事件
        for f in extracted_files:
            if os.path.isdir(f):
                # 如果是目录
                self.log(f'找到目录: {os.path.basename(f)}')
            else:
                file_queue.put(f)  # 将提取的文件放入文件队列
                self.log(f'添加待处理文件: {os.path.basename(f)}')
        return True

    def extract_file(self, file_path):
        # 解压文件
        temp_dir = self.generate_temp_dir()  # 生成临时目录
        try:
            passwords = [None] + self.passwords  # 密码列表，包括无密码
            for pwd in passwords:
                # 尝试不同的密码
                self.log(f'尝试密码: {pwd}'if pwd else'尝试无密码解压')
                result = self.run_7z_command(file_path, temp_dir, pwd)  # 执行7z命令
                if result.returncode == 0:
                    # 如果解压成功
                    self.log('解压成功')
                    break
            else:
                # 如果所有密码尝试都失败
                self.log(f'所有密码尝试失败')
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)  # 删除临时目录
                return []
            extracted_files = []
            for root, dirs, files in os.walk(temp_dir):
                # 遍历临时目录
                self.log(f'当前目录: {dirs+files}')
                for name in dirs + files:
                    try:
                        filename = os.path.join('.', name)  # 目标文件路径
                        os.rename(os.path.join(root, name), filename)
                        extracted_files.append(filename)
                    except Exception as e:
                        self.log(f'移动文件失败: {str(e)}')
            shutil.rmtree(temp_dir)  # 删除临时目录
            return extracted_files
        except Exception as e:
            self.log(f'解压错误: {str(e)}')
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)  # 删除临时目录
            return []

    def run_7z_command(self, file_path, temp_dir, password=None):
        # 执行7z解压命令
        seven_zip_path = self.get_7z_path()  # 获取7z路径
        cmd = [seven_zip_path, 'x', '-y', file_path, f'-o{temp_dir}']  # 7z命令
        if not password:
            password = ''
        cmd.insert(5, '-aoa')  # 添加覆盖所有文件的参数
        cmd.insert(5, f'-p{password}')  # 添加密码参数
        # 执行7z命令，隐藏窗口
        return subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW)

if __name__ == '__main__':
    # 创建应用实例
    app = AutoUnzipApp()
    # 绑定状态更新事件
    app.window.bind('<<UpdateStatus>>', lambda e: app.update_status())
    # 启动主循环
    app.window.mainloop()