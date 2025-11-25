import locale# locale: 用于处理本地化设置，确保日期时间格式正确显示
import os# os: 提供操作系统接口，用于文件和路径操作
import sys# sys: 提供对Python解释器的访问，用于获取程序路径等信息
import subprocess# subprocess: 用于创建子进程执行外部命令（如7-Zip）
import shutil# shutil: 提供高级文件操作，用于复制、移动和删除文件/目录
import threading# threading: 提供线程支持，实现多线程并行处理
import queue# queue: 提供线程安全的队列实现，用于任务管理
import tkinter as tk# tkinter: GUI库，用于创建图形用户界面
import random# random: 提供随机数生成功能，用于生成临时文件名
import string# string: 提供字符串常量和处理函数
import re# re: 提供正则表达式支持，用于文件名匹配（如分卷压缩文件）
from tkinter import scrolledtext# scrolledtext: 提供带滚动条的文本组件，用于显示日志信息


class AutoUnzipApp:
    """
    自动解压应用类
    
    该类提供了一个图形界面的自动解压工具，支持多线程解压7z、rar、zip等格式的压缩文件，
    并能尝试使用密码列表解密受密码保护的文件。同时支持分卷压缩文件的识别和处理。
    
    主要功能特性：
    - 多线程并行解压，提高处理效率
    - 支持多种压缩格式：7z、rar、zip及其分卷格式
    - 自动尝试密码列表中的密码进行解密
    - 自动识别和处理分卷压缩文件
    - 自动递归解压嵌套的压缩文件
    - 提供图形界面显示解压进度和日志信息
    """
    def __init__(self):
        # 初始化线程安全的队列，用于存储待显示的日志消息
        self.log_queue = queue.Queue()
        # 初始化线程安全的任务队列，用于存储待处理的文件路径
        self.task_queue = queue.Queue()
        # 创建线程锁，用于保护共享资源的访问
        self.lock = threading.Lock()
        
        # 创建Tkinter主窗口
        self.window = tk.Tk()
        # 设置窗口标题
        self.window.title('auto_unzip')
        # 设置窗口初始大小为800x600像素
        self.window.geometry('800x600')
        
        # 创建可滚动的文本区域，用于显示日志信息
        self.log_area = scrolledtext.ScrolledText(self.window, wrap=tk.WORD)
        # 将日志区域填充整个窗口空间
        self.log_area.pack(expand=True, fill='both')
        
        # 创建状态栏框架
        self.status_frame = tk.Frame(self.window)
        # 将状态栏框架填充水平空间
        self.status_frame.pack(fill=tk.X)
        
        # 创建状态栏标签，显示解压次数
        self.status_label = tk.Label(self.status_frame, text='解压次数: 0')
        # 将状态栏标签放置在左侧
        self.status_label.pack(side=tk.LEFT)
        
        # 输出开源信息到日志
        self.log('本工具已开源在https://github.com/liaoran2002/auto_unzip')
        
        # 初始化活跃任务计数器
        self.active_tasks = 0
        # 初始化已处理文件计数器
        self.process_count = 0
        # 设置单个文件的最大尝试次数，防止无限递归
        self.max_attempts = 10
        # 设置最大工作线程数量
        self.max_workers = 4
        # 设置临时目录前缀
        self.temp_dir_prefix = 'tmp_'
        # 定义支持的压缩文件扩展名集合
        self.compressed_exts = {'7z', 'rar', 'zip'}
        # 编译正则表达式，用于匹配分卷压缩文件的命名模式
        self.split_file_pattern = re.compile(r'(\.part\d+\.rar|\.r\d{2,}|\.7z\.\d{3,}|\.\d{3,})$', re.IGNORECASE)
        
        # 构建密码文件的完整路径
        self.passwords_file_path = os.path.join(self.get_exe_dir(), 'passwords.txt')
        # 初始化密码列表
        self.passwords = []
        
        # 检查密码文件是否存在
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
        
        # 创建工作线程池，启动指定数量的工作线程
        for _ in range(self.max_workers):
            threading.Thread(target=self.process_worker, daemon=True).start()
        
        # 设置定时器，每100毫秒检查一次日志队列
        self.window.after(100, self.check_log_queue)
        
        # 检查命令行参数，如果有参数则将其作为文件路径添加到任务队列
        if len(sys.argv) > 1:
            for file_path in sys.argv[1:]:
                self.add_task(file_path)

    def get_7z_path(self):
        """
        获取7z可执行文件的路径
        
        返回值：
            str: 7z可执行文件的完整路径
            
        该方法根据程序运行环境判断7z可执行文件的位置：
        - 如果程序是打包后的可执行文件（frozen），则从sys._MEIPASS获取路径
        - 否则，从当前脚本所在目录获取7z可执行文件路径
        
        这种设计确保了无论是在开发环境还是打包后的环境中，都能正确找到7z可执行文件。
        """
        
        # 检查程序是否被打包（如使用PyInstaller打包）
        if getattr(sys, 'frozen', False):
            # 当程序被打包时，使用sys._MEIPASS获取临时提取目录
            base_path = sys._MEIPASS
        else:
            # 在开发环境中，获取当前脚本所在目录的绝对路径
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        # 构建7z可执行文件的完整路径
        # 注意：这里直接使用'7z'作为文件名，在不同平台上可能需要调整
        # Windows系统上会自动查找'7z.exe'，而在Unix/Linux系统上则直接使用'7z'
        return os.path.join(base_path, '7z')

    def get_exe_dir(self):
        """
        获取可执行文件所在目录的路径
        
        返回值：
            str: 可执行文件所在目录的完整路径
            
        该方法根据程序运行环境确定程序所在目录：
        - 如果程序是打包后的可执行文件（frozen），则使用sys.executable获取可执行文件路径
        - 否则，使用当前脚本文件的绝对路径
        
        此方法用于确定配置文件、密码文件等资源文件的位置。
        """
        if getattr(sys, 'frozen', False):
            exe_path = sys.executable
        else:
            exe_path = os.path.abspath(__file__)
        return os.path.dirname(exe_path)

    def generate_temp_dir(self):
        """
        生成临时目录名称
        
        返回值：
            str: 临时目录名称
            
        该方法生成一个随机的临时目录名称，由固定前缀和随机字符串组成。
        使用字母和数字的组合生成5个字符的随机字符串，确保临时目录名称的唯一性，
        避免在多线程环境下产生目录冲突。
        """
        chars = string.ascii_lowercase + string.digits
        return self.temp_dir_prefix + ''.join(random.choices(chars, k=5))

    def check_log_queue(self):
        """
        检查日志队列并更新GUI显示
        
        该方法是一个定时执行的函数，每100毫秒执行一次，用于：
        1. 检查日志消息队列是否有新消息
        2. 如果有新消息，将其添加到日志显示区域
        3. 自动滚动到最新消息位置
        4. 设置下一次检查的定时器
        
        这种设计确保了在多线程环境下，日志消息能够安全地从工作线程传递到UI线程并显示。
        """
        if not self.log_queue.empty():
            msg = self.log_queue.get()
            self.log_area.insert(tk.END, msg + '\n')
            self.log_area.see(tk.END)
        self.window.after(100, self.check_log_queue)

    def log(self, message):
        """
        记录日志消息
        
        参数：
            message (str): 要记录的日志消息
            
        该方法将日志消息放入线程安全的队列中，供UI线程的check_log_queue方法处理。
        这确保了在多线程环境中，所有线程都能安全地输出日志，而不会导致UI更新冲突。
        """
        # 导入datetime模块，获取当前时间戳
        import datetime
        # 格式化当前时间为字符串（年-月-日 时:分:秒）
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # 构建完整的日志消息，包含时间戳和原始消息
        log_message = f'[{timestamp}] {message}'
        # 将格式化后的日志消息添加到日志队列中，确保线程安全
        self.log_queue.put(log_message)

    def update_status(self):
        """
        更新状态栏显示
        
        该方法更新GUI状态栏，显示当前的解压次数和活跃任务数量。
        通常通过事件绑定（<<UpdateStatus>>）触发，确保UI及时反映任务进度。
        """
        # 构建状态栏文本，显示已解压文件次数和当前活跃任务数量
        status_text = f'已解压次数: {self.process_count} \t 活跃任务数: {self.active_tasks}'
        # 更新状态栏标签的文本内容，自动触发UI刷新
        self.status_label.config(text=status_text)

    def add_task(self, file_path):
        """
        添加解压任务到任务队列
        
        参数：
            file_path (str): 要解压的文件路径
            
        该方法将文件路径转换为绝对路径，验证文件存在性，
        然后将文件添加到任务队列中，并增加活跃任务计数。
        使用线程锁确保任务添加和计数更新的原子性。
        """
        # 将相对路径转换为绝对路径，确保文件引用的一致性
        abs_path = os.path.abspath(file_path)
        
        # 验证文件是否存在，避免处理不存在的文件
        if not os.path.exists(abs_path):
            self.log(f'文件不存在: {abs_path}')
            return
        
        # 使用线程锁保护共享资源，确保多线程环境下的操作原子性
        with self.lock:
            # 将文件路径添加到任务队列，等待工作线程处理
            self.task_queue.put(abs_path)
            
            # 记录任务添加日志，使用绝对路径便于追踪
            self.log(f'已添加任务: {abs_path}')
            
            # 增加活跃任务计数，用于状态跟踪
            self.active_tasks += 1

    def process_worker(self):
        """
        工作线程处理函数
        
        该方法是一个无限循环，作为工作线程的主函数，负责：
        1. 从任务队列中获取文件路径
        2. 调用process_single_file方法处理文件
        3. 任务完成后减少活跃任务计数
        4. 当所有任务完成时，延迟关闭窗口
        
        该方法使用try-finally结构确保即使处理过程中出现异常，
        也能正确标记任务完成并更新状态。
        """
        # 无限循环，持续处理任务队列中的文件
        while True:
            # 从任务队列中获取一个文件路径（如果队列为空则阻塞）
            current_path = self.task_queue.get()
            try:
                # 调用process_single_file方法处理当前文件
                self.process_single_file(current_path)
            finally:
                # 无论处理是否成功，都标记任务完成
                self.task_queue.task_done()
                # 使用线程锁保护共享资源的访问
                with self.lock:
                    # 减少活跃任务计数
                    self.active_tasks -= 1
                    # 当没有活跃任务时，延迟1秒后关闭窗口
                    if self.active_tasks == 0:
                        self.window.after(1000, self.window.destroy)

    def is_compressed_file(self, filename):
        """
        判断文件是否为压缩文件
        
        参数：
            filename (str): 要检查的文件名或路径
            
        返回值：
            bool: 如果是压缩文件返回True，否则返回False
            
        该方法通过两种方式判断文件是否为压缩文件：
        1. 检查文件扩展名是否在支持的压缩格式列表中（7z、rar、zip）
        2. 使用正则表达式检查文件名是否匹配分卷压缩文件模式
        
        这样可以识别常见的压缩文件及其分卷格式。
        """
        # 获取文件名部分（不包含路径）
        base_name = os.path.basename(filename)
        # 获取文件扩展名并转换为小写
        ext = base_name.split('.')[-1].lower()
        # 检查扩展名是否在支持的压缩格式列表中
        if ext in self.compressed_exts:
            return True
        # 检查文件名是否匹配分卷压缩文件的模式
        return self.split_file_pattern.search(base_name) is not None

    def process_single_file(self, initial_path):
        """
        处理单个文件的主函数
        
        参数：
            initial_path (str): 要处理的初始文件路径
            
        该方法是文件处理的核心，负责：
        1. 创建一个本地队列用于递归处理解压出的文件
        2. 设置最大尝试次数，避免无限递归
        3. 循环处理队列中的每个文件
        4. 对于压缩文件，调用extract_file方法解压
        5. 对于非压缩文件，尝试作为压缩文件处理
        6. 处理解压结果并更新计数
        
        这种设计支持自动递归解压嵌套的压缩文件。
        """
        # 创建本地队列，用于存储待处理的文件路径
        file_queue = queue.Queue()
        # 将初始文件路径加入队列
        file_queue.put(initial_path)
        # 初始化本地计数器，用于限制最大处理次数
        local_count = 0
        
        # 循环处理队列中的文件，直到队列为空或达到最大尝试次数
        while not file_queue.empty() and local_count < self.max_attempts:
            # 从队列中获取下一个要处理的文件路径
            current_path = file_queue.get()
            
            # 检查是否为目录
            if os.path.isdir(current_path):
                self.log(f'找到目录: {os.path.basename(current_path)}')
                # 跳过目录，继续处理队列中的下一个文件
                continue
            
            # 获取原始文件名
            original_name = os.path.basename(current_path)
            # 初始化处理成功标志
            success = False
            
            # 判断是否为压缩文件
            if self.is_compressed_file(current_path):
                self.log(f'检测到压缩文件: {original_name}')
                # 调用extract_file方法解压文件
                extracted_files = self.extract_file(current_path)
                # 处理解压结果
                success = self.handle_extraction_result(current_path, extracted_files, file_queue)
            else:
                # 非压缩文件，尝试作为压缩文件处理
                self.log(f'检测到非压缩文件: {original_name}，尝试重命名为压缩文件格式')
                success = self.process_as_non_compressed(current_path, file_queue)
            
            # 根据处理结果更新计数
            if success:
                local_count += 1
            else:
                self.log(f'文件无法解压: {current_path}')
        
        # 记录处理完成信息
        self.log(f'文件 {initial_path} 处理完成：达到最大次数或找到目录')

    def process_as_non_compressed(self, current_path, file_queue):
        original_name = os.path.basename(current_path)
        for ext in self.compressed_exts:
            temp_file = f'{current_path}.{ext}'
            with self.lock:
                if not os.path.exists(current_path):
                    self.log(f'文件已被其他进程处理: {original_name}')
                    continue
                try:
                    os.rename(current_path, temp_file)
                except Exception as e:
                    self.log(f'重命名失败: {str(e)}')
                    continue
            self.log(f'尝试解压: {temp_file}')
            extracted_files = self.extract_file(temp_file)
            if extracted_files:
                self.handle_extraction_result(temp_file, extracted_files, file_queue)
                return True
            else:
                with self.lock:
                    if os.path.exists(temp_file):
                        os.rename(temp_file, current_path)
        return False

    def find_split_files(self, main_file):
        """
        查找与指定压缩文件相关的所有分卷文件
        
        参数：
            main_file (str): 主要压缩文件的路径
            
        返回值：
            list: 包含所有分卷文件路径的列表
            
        该方法通过以下步骤查找分卷文件：
        1. 解析主文件名，提取基本名称和数字标识
        2. 使用正则表达式查找匹配的分卷文件
        3. 按照正确的顺序排序分卷文件
        4. 返回完整的分卷文件列表
        """
        # 获取主文件的文件名部分（不含路径）
        base_name = os.path.basename(main_file)
        # 使用正则表达式检查文件名是否匹配分卷文件模式
        match = self.split_file_pattern.search(base_name)
        # 如果不匹配分卷模式，则没有分卷文件，返回空列表
        if not match:
            return []
        # 移除非分卷标识的基础文件名模式
        base_pattern = self.split_file_pattern.sub('', base_name)
        # 获取主文件所在的目录路径
        dir_path = os.path.dirname(main_file)
        # 创建列表用于存储找到的分卷文件
        split_files = []
        # 遍历目录中的所有文件
        for f in os.listdir(dir_path):
            # 检查文件名是否以基础模式开头，并且不是主文件本身
            if f.startswith(base_pattern) and f != base_name:
                # 构建完整的文件路径
                full_path = os.path.join(dir_path, f)
                # 再次使用正则表达式确认是分卷文件
                if self.split_file_pattern.search(f):
                    # 添加到分卷文件列表
                    split_files.append(full_path)
        # 对分卷文件列表进行排序，确保按正确顺序处理
        split_files.sort()
        # 返回排序后的分卷文件列表
        return split_files

    def handle_extraction_result(self, src_file, extracted_files, file_queue):
        """
        处理解压结果
        
        参数：
            src_file (str): 原始压缩文件路径
            extracted_files (list): 解压出的文件列表
            file_queue (queue.Queue): 文件队列，用于添加解压出的待处理文件
            
        该方法负责处理解压完成后的后续操作：
        1. 删除原始压缩文件及分卷文件
        2. 将解压出的文件添加到队列中进行进一步处理
        3. 更新解压计数
        4. 记录操作日志
        
        支持递归处理解压出的嵌套压缩文件。
        """
        # 检查是否成功解压出文件，如果解压失败则直接返回False
        if not extracted_files:
            return False
        
        # 查找与源文件相关的分卷压缩文件
        split_files = self.find_split_files(src_file)
        if split_files:
            self.log(f'发现分卷文件组：共 {len(split_files)} 个')
        
        # 创建集合用于跟踪已删除的文件，避免重复删除
        deleted_files = set()
        
        # 处理所有分卷文件
        for f in split_files:
            # 使用线程锁保护文件删除操作，确保多线程环境下的安全
            with self.lock:
                try:
                    # 检查文件是否存在且未被删除过
                    if os.path.exists(f) and f not in deleted_files:
                        # 删除分卷文件
                        os.remove(f)
                        # 将已删除文件添加到已删除集合
                        deleted_files.add(f)
                        # 记录删除操作日志
                        self.log(f'删除分卷文件: {os.path.basename(f)}')
                except Exception as e:
                    # 记录删除分卷文件失败的错误信息
                    self.log(f'删除分卷失败: {str(e)}')
        
        # 删除原始压缩文件
        with self.lock:
            try:
                # 检查文件是否存在且未被删除过
                if os.path.exists(src_file) and src_file not in deleted_files:
                    # 删除原始文件
                    os.remove(src_file)
                    # 将已删除文件添加到已删除集合
                    deleted_files.add(src_file)
                    # 记录删除操作日志
                    self.log(f'删除原文件: {os.path.basename(src_file)}')
            except Exception as e:
                # 记录删除原始文件失败的错误信息
                self.log(f'删除原文件失败: {str(e)}')
        
        # 增加已处理文件计数
        self.process_count += 1
        # 触发UI更新事件，刷新状态栏显示
        self.window.event_generate('<<UpdateStatus>>')
        
        # 处理解压出的所有文件
        for f in extracted_files:
            # 如果是目录，只记录日志
            if os.path.isdir(f):
                self.log(f'找到目录: {os.path.basename(f)}')
            else:
                # 如果是文件，将其添加到文件队列中，以便递归处理（处理嵌套的压缩文件）
                file_queue.put(f)
                self.log(f'添加待处理文件: {os.path.basename(f)}')
        
        # 返回处理成功标志
        return True

    def extract_file(self, file_path):
        """
        解压文件的核心方法
        
        参数：
            file_path (str): 要解压的文件路径
            
        返回值：
            list: 成功解压的文件列表
        
        该方法使用7-Zip命令行工具执行解压操作，主要功能包括：
        1. 构建解压命令行参数
        2. 尝试使用密码列表中的密码进行解压
        3. 处理解压过程中的错误
        4. 记录解压日志
        5. 返回解压结果
        
        支持处理多种压缩格式和密码保护的文件。
        """
        # 生成临时目录用于存放解压文件
        temp_dir = self.generate_temp_dir()
        try:
            # 构建密码尝试列表，首先尝试无密码
            passwords = [None] + self.passwords
            # 遍历密码列表尝试解压
            for pwd in passwords:
                # 记录当前尝试的密码信息
                self.log(f'尝试密码: {pwd}' if pwd else '尝试无密码解压')
                # 调用run_7z_command执行实际的解压操作
                result = self.run_7z_command(file_path, temp_dir, pwd)
                # 检查解压是否成功（返回码为0表示成功）
                if result.returncode == 0:
                    self.log('解压成功')
                    break
            else:
                # 所有密码尝试都失败
                self.log(f'所有密码尝试失败')
                # 清理临时目录
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                # 返回空列表表示解压失败
                return []
            
            # 初始化解压文件列表
            extracted_files = []
            # 遍历临时目录中的所有文件和子目录
            for root, dirs, files in os.walk(temp_dir):
                # 记录当前目录内容
                self.log(f'当前目录内容: {dirs + files}')
                # 处理每个文件和目录
                for name in dirs + files:
                    try:
                        # 构建目标文件路径（当前目录）
                        filename = os.path.join('.', name)
                        # 将解压出的文件移动到当前目录
                        os.rename(os.path.join(root, name), filename)
                        # 将移动成功的文件添加到结果列表
                        extracted_files.append(filename)
                    except Exception as e:
                        # 处理重命名失败的情况
                        self.log(f'重命名文件失败: {str(e)}')
            # 解压完成后删除临时目录
            shutil.rmtree(temp_dir)
            return extracted_files
        except Exception as e:
            self.log(f'解压错误: {str(e)}')
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return []

    def run_7z_command(self, file_path, temp_dir, password=''):
        seven_zip_path = self.get_7z_path()
        cmd = [seven_zip_path, 'x', '-y', file_path, f'-o{temp_dir}']
        cmd.insert(5, '-aoa')
        cmd.insert(5, f'-p{password}')
        return subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                              creationflags=subprocess.CREATE_NO_WINDOW)


def main():
    """
    程序主入口函数
    
    该函数负责初始化应用程序并启动UI事件循环。
    具体功能包括：
    1. 创建AutoUnzipApp实例，这会初始化整个应用程序
    2. 启动Tkinter的主事件循环，使GUI界面保持响应状态
    """
    # 创建AutoUnzipApp实例，初始化应用程序的所有组件
    # 包括UI界面、工作线程、各种队列和状态变量
    AutoUnzipApp()
    
    # 启动Tkinter的主事件循环
    # 这是GUI应用程序的核心，会持续监听用户输入和事件
    # 直到用户关闭窗口或程序退出
    tk.mainloop()


# 程序入口点检查
# 当此脚本作为主程序运行时（而不是作为模块导入时）
# 执行main()函数启动应用程序
if __name__ == '__main__':
    main()