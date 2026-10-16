# -*- coding: utf-8 -*-
"""auto_unzip —— 自动解压工具(带图形界面)

相对早期版本的关键修正:

1. 解压产物落在**压缩包所在目录**,不再倒进进程的当前工作目录。
2. 先解压到临时目录,再搬到目标目录;**搬运有任何失败就保留原包**,
   杜绝"文件没搬出来、原包却被删了"的静默丢数据。
3. 用**文件头(magic bytes)**判断真实格式,不再靠扩展名猜:
   - 放过 docx/xlsx/pptx/jar/apk/epub 这类"内里就是 ZIP"的文档容器;
   - 放过 video.mp4.001、keygen.001 这类伪装成 .001 的普通文件。
4. 分卷按"词干 + 分卷类型"**精确分组**,不再按文件名前缀误删无关文件。
5. 工作线程加异常兜底,避免一个未捕获异常把线程打死、
   任务被丢弃、窗口永远不关闭。
6. 临时目录建在压缩包同盘,保证搬运不跨卷。
7. 解压结果里若出现**子目录中的压缩包**,也会继续处理。
8. 同名文件自动改名(_1/_2...),不再覆盖也不再把冲突当作失败。
9. 启动时自检 7z 引擎;RAR 需要完整版 7-Zip,缺失时明确提示而不是静默失败。
"""

import locale
import os
import re
import shutil
import subprocess
import sys
import queue
import tempfile
import threading
import time
import tkinter as tk
from tkinter import scrolledtext, ttk
import uuid
import zipfile
from tkinter import scrolledtext

# ---- 压缩格式的文件头(magic bytes) --------------------------------------
# 判断"这个文件到底是什么",而不是看名字。顺序有意义:先匹配到的为准。
MAGIC_SIGNATURES = (
    (b'7z\xbc\xaf\x27\x1c', '7z'),
    (b'PK\x03\x04', 'zip'),            # 常规 zip / docx / jar ...
    (b'PK\x05\x06', 'zip'),            # 空 zip
    (b'PK\x07\x08', 'zip'),            # zip 续卷
    (b'Rar!\x1a\x07\x01\x00', 'rar'),  # RAR 5.x(更长,要先比)
    (b'Rar!\x1a\x07\x00', 'rar'),      # RAR 4.x
    (b'\x1f\x8b', 'gzip'),
    (b'BZh', 'bzip2'),
    (b'\xfd7zXZ\x00', 'xz'),
    (b'\x28\xb5\x2f\xfd', 'zstd'),
    (b'MSCF', 'cab'),
)

# 这些"内里其实是 ZIP"的文档容器,不该被解压后删掉原文件。
CONTAINER_MARKERS = ('[Content_Types].xml', 'META-INF/MANIFEST.MF',
                     'AndroidManifest.xml', 'mimetype')
CONTAINER_DIRS = ('word/', 'xl/', 'ppt/', 'META-INF/', 'OEBPS/')

# 进度行里文件名一栏的宽度(像素)与最大字符数(超出则中间省略)
PROGRESS_NAME_PX = 250
PROGRESS_NAME_CHARS = 40

# 图种(图片里附带压缩包)常见的文件类型。
POLYGLOT_EXTS = {
    '.jpg', '.jpeg', '.jpe', '.jfif', '.png', '.gif', '.bmp',
    '.webp', '.tif', '.tiff', '.ico', '.heic', '.avif',
    # 少数情况会塞进这些里
    '.mp3', '.mp4', '.wav', '.flac', '.mkv', '.avi',
    '.pdf', '.doc', '.xls', '.ppt',
}

# 分卷命名规则:(正则, 类型标记)。类型标记不同就不算同一组。
VOLUME_PATTERNS = (
    (re.compile(r'\.part(\d+)\.rar$', re.IGNORECASE), 'rar_part'),
    (re.compile(r'\.r(\d{2,})$', re.IGNORECASE), 'rar_old'),
    (re.compile(r'\.7z\.(\d{3,})$', re.IGNORECASE), '7z_split'),
    (re.compile(r'\.zip\.(\d{3,})$', re.IGNORECASE), 'zip_split'),
    (re.compile(r'\.z(\d{2,})$', re.IGNORECASE), 'zip_z'),
    (re.compile(r'\.(\d{3,})$', re.IGNORECASE), 'generic_split'),
)


def _read_head(path, size=512):
    """读文件头,读不到就返回空字节。"""
    try:
        with open(path, 'rb') as f:
            return f.read(size)
    except OSError:
        return b''


def detect_format(path):
    """按文件头识别真实格式。返回 '7z'/'zip'/'rar'/... 或 None。"""
    head = _read_head(path)
    if not head:
        return None
    for sig, kind in MAGIC_SIGNATURES:
        if head.startswith(sig):
            return kind
    # tar 的 'ustar' 在偏移 257
    if len(head) > 262 and head[257:262] == b'ustar':
        return 'tar'
    return None


def is_document_container(path):
    """判断这个 ZIP 是不是 Office/JAR/APK/EPUB 之类的文档容器。"""
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
    except Exception:
        return False
    for marker in CONTAINER_MARKERS:
        if marker in names:
            return True
    for name in names:
        if name.startswith(CONTAINER_DIRS):
            return True
    return False


def volume_info(name):
    """拆出 (词干, 分卷类型, 序号);不是分卷则返回 None。"""
    for pattern, kind in VOLUME_PATTERNS:
        m = pattern.search(name)
        if m:
            base = name[:m.start()]
            try:
                num = int(m.group(1))
            except (TypeError, ValueError):
                num = 0
            return base, kind, num
    return None


class AutoUnzipApp:
    def __init__(self):
        # 线程间传递日志
        self.log_queue = queue.Queue()
        # 待处理的初始文件
        self.task_queue = queue.Queue()
        # 保护共享状态(任务计数/解压计数)
        self.lock = threading.Lock()
        # 保护"往目标目录搬文件"这一动作:多个线程可能同时往同一目录搬,
        # 必须保证"选名字 + 搬过去"是原子的,否则会互相覆盖丢文件。
        self.move_lock = threading.RLock()

        # ---- 可调参数 ----
        self.max_attempts = 20            # 单个输入最多连续解压多少层
        self.max_workers = 4              # 工作线程数
        self.temp_dir_prefix = 'tmp_'
        self.extract_timeout = 1800       # 单次 7z 调用的超时(秒)
        self.protect_documents = True     # 保护 docx/xlsx 等文档容器
        self.scan_polyglot = True         # 识别"图种"(图片里附带的压缩包)
        self.scan_timeout = 60            # 图种探测(7z l)的超时(秒)
        self.close_delay_ms = 1500        # 全部完成后延迟关窗
        self.show_progress = True         # 显示解压进度条
        self.progress_interval_ms = 120   # 进度条刷新间隔(毫秒)
        self.progress_stall_s = 300       # 多久没进度就认为卡住(秒)

        # ---- 状态 ----
        self.active_tasks = 0
        self.process_count = 0
        self.passwords = []
        self.seven_zip = None
        self.seven_zip_desc = ''
        self.rar_capable = False
        self._tasks_added = False
        self._stay_open = False
        self._close_pending = False
        # 进度信息(worker 写,主线程读)
        # 同时可能有多个压缩包在解压,所以用 dict 保存每一条进度:
        #   task_id -> {'percent':int, 'label':str, 'done':bool, 'updated':float}
        self.progress_lock = threading.Lock()
        self._progress = {}
        self._progress_seq = 0

        # ---- 界面 ----
        self.window = tk.Tk()
        self.window.title('auto_unzip')
        self.window.geometry('820x600')

        # ---- 进度区(每个正在解压的包一行:左文件名 | 中进度条 | 右百分比)----
        # 状态栏先 pack(占最底部),进度区再 pack(落在状态栏上方),
        # 这样进度区在日志和状态栏之间,比较醒目。
        self.status_frame = tk.Frame(self.window)
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label = tk.Label(self.status_frame, text='解压次数: 0')
        self.status_label.pack(side=tk.LEFT)

        self.progress_outer = tk.Frame(self.window)
        self.progress_outer.pack(side=tk.BOTTOM, fill=tk.X)
        self.progress_frame = tk.Frame(self.progress_outer)
        self.progress_frame.pack(fill=tk.X)
        # task_id -> {'frame','name','bar','pct'}
        self.progress_rows = {}

        # 日志区占满剩余空间(最后 pack,所以放最上面)
        self.log_area = scrolledtext.ScrolledText(self.window, wrap=tk.WORD)
        self.log_area.pack(expand=True, fill='both')

        self.log('本工具已开源在https://github.com/liaoran2002/auto_unzip')

        # ---- 定位 7z 可执行文件 ----
        self._locate_7z()

        # ---- 加载密码表 ----
        self._load_passwords()

        # ---- 起工作线程 ----
        if self.seven_zip:
            for _ in range(self.max_workers):
                threading.Thread(target=self.process_worker, daemon=True).start()

        # ---- 主线程定时器:刷日志 + 刷状态 + 判断收工 ----
        self.window.after(100, self.check_log_queue)

        # ---- 命令行参数当任务 ----
        if len(sys.argv) > 1:
            for file_path in sys.argv[1:]:
                self.add_task(file_path)
        else:
            self._stay_open = True
            self.log('请把压缩包拖到 auto_unzip 图标上使用;')
            self.log('或先用命令行: auto_unzip.exe 某个压缩包.zip')
            self.log('(没有收到文件,窗口保持打开,关掉即可。)')

    # ------------------------------------------------------------------
    # 7z 可执行文件
    # ------------------------------------------------------------------
    def get_exe_dir(self):
        """程序(exe 或脚本)所在目录。"""
        if getattr(sys, 'frozen', False):
            return os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(os.path.abspath(__file__))

    def _candidate_7z_paths(self):
        """按优先级列出可能的 7z 可执行文件。

        首选随包的 7z.exe(完整版,配合同目录 7z.dll 支持 RAR);
        其次才找只支持部分格式的 7za.exe,最后退回系统安装的 7-Zip。
        """
        here = self.get_exe_dir()
        bases = [here]
        if getattr(sys, 'frozen', False):
            # PyInstaller 单文件模式把 --add-data 解到 _MEIPASS
            meipass = getattr(sys, '_MEIPASS', None)
            if meipass:
                bases.insert(0, meipass)
        cands = []
        for base in bases:
            # 完整版(要 7z.dll 陪着,支持 RAR)
            cands.append(os.path.join(base, '7z.exe'))
            # 只支持部分格式的独立版(不需要 dll,除 RAR 外也够用)
            cands.append(os.path.join(base, '7za.exe'))
        # 保留对旧目录结构的兼容
        for base in bases:
            cands.append(os.path.join(base, '7z2604-extra', 'x64', '7za.exe'))
            cands.append(os.path.join(base, '7z2604-extra', '7za.exe'))
        # 环境里已有的
        for name in ('7z.exe', '7za.exe', '7za'):
            found = shutil.which(name)
            if found:
                cands.append(found)
        # 最后才用系统安装的(可能带 RAR)
        for key in ('ProgramFiles', 'ProgramFiles(x86)'):
            root = os.environ.get(key)
            if root:
                cands.append(os.path.join(root, '7-Zip', '7z.exe'))
                cands.append(os.path.join(root, '7-Zip', '7za.exe'))
        return cands

    def _locate_7z(self):
        first_usable = None
        for cand in self._candidate_7z_paths():
            if not os.path.isfile(cand):
                continue
            desc, formats = self._probe_7z(cand)
            if desc is None:
                continue
            rar = self._supports(formats, 'rar', 'rar5')
            if first_usable is None:
                first_usable = (cand, desc, rar)
            if rar:
                # 优先选支持 RAR 的引擎(完整版 7z.exe + 7z.dll)
                self._use_engine(cand, desc, True)
                return
        if first_usable:
            cand, desc, rar = first_usable
            self._use_engine(cand, desc, rar)
            return
        self._stay_open = True
        self.log('=' * 60)
        self.log('错误: 没找到可用的 7z 解压引擎,无法工作。')
        self.log('请把 7z.exe 和 7z.dll 一起放在本程序同目录下。')
        self.log('=' * 60)

    def _use_engine(self, path, desc, rar):
        self.seven_zip = path
        self.seven_zip_desc = desc
        self.rar_capable = rar
        self.log(f'使用解压引擎: {desc}')
        self.log(f'  路径: {path}')
        if rar:
            self.log('  该引擎支持 RAR。')
        else:
            self.log('  提示: 该引擎不支持 RAR。若要解压 RAR,')
            self.log('        请把完整版 7-Zip 的 7z.exe 和 7z.dll 一起放到本程序同目录。')

    @staticmethod
    def _probe_7z(exe):
        """确认 exe 可用,并返回 (版本描述, 支持格式集合)。

        注意:必须用 'i' 命令才能拿到格式表;不带参数只会打印帮助文本,
        那样永远探测不到 RAR(早期版本就是这样误报"不支持 RAR"的)。
        """
        kwargs = {}
        flag = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        if flag:
            kwargs['creationflags'] = flag

        def run(args):
            try:
                r = subprocess.run([exe] + args, stdout=subprocess.PIPE,
                                   stderr=subprocess.STDOUT, timeout=25, **kwargs)
            except Exception:
                return None
            return (r.stdout or b'').decode('utf-8', 'replace')

        text = run(['i'])
        if text is None:
            text = run([])          # 极老的版本可能不认 i
        if not text:
            return None, set()

        desc = ''
        for line in text.splitlines():
            if '7-Zip' in line:
                desc = line.strip()
                break
        if not desc:
            return None, set()

        # 从格式表里解析出该引擎真正支持的格式名。
        # 行形如:  0 ...F...  Rar      rar r00       R a r ! 1A 07 00
        formats = set()
        for line in text.splitlines():
            m = re.match(r'\s*(?:\d+\s+)?[CFKMGOPXELHBW\s.]{2,}\s+(\S+)\s+(\S+)', line)
            if not m:
                continue
            name = m.group(1)
            ext_field = m.group(2)
            if not name or name.lower() in ('formats:', 'codecs:', 'hashers:'):
                continue
            formats.add(name.lower())
            for ext in ext_field.split():
                formats.add(ext.lower().lstrip('.'))
        return desc, formats

    @staticmethod
    def _supports(formats, *names):
        """判断引擎是否支持某个格式(按格式名或扩展名)。"""
        low = {f.lower().lstrip('.') for f in (formats or ())}
        return any(n.lower() in low for n in names)

    # ------------------------------------------------------------------
    # 密码表
    # ------------------------------------------------------------------
    def _load_passwords(self):
        self.passwords_file_path = os.path.join(self.get_exe_dir(), 'passwords.txt')
        if not os.path.exists(self.passwords_file_path):
            self._create_password_files()
            return

        # 去重后的候选编码
        encodings = []
        for enc in ('utf-8', locale.getpreferredencoding(False), 'gbk'):
            if enc and enc.lower() not in [e.lower() for e in encodings]:
                encodings.append(enc)

        for encoding in encodings:
            try:
                with open(self.passwords_file_path, 'r', encoding=encoding) as f:
                    raw = f.readlines()
            except UnicodeDecodeError:
                continue
            except Exception as e:
                self.log(f'读取密码文件出错({encoding}): {e}')
                break
            # 忽略空行和注释(README 里承诺过要忽略注释)
            self.passwords = [
                line.strip() for line in raw
                if line.strip() and not line.lstrip().startswith(('#', ';'))
            ]
            self.log(f'已用 {encoding} 加载 {len(self.passwords)} 个密码')
            return
        else:
            self.passwords = []
            self.log('所有编码都读不了密码文件,已忽略')

    def _create_password_files(self):
        try:
            with open(self.passwords_file_path, 'w', encoding='utf-8') as f:
                f.write('')
            example = os.path.join(self.get_exe_dir(), 'passwords_example.txt')
            with open(example, 'w', encoding='utf-8') as f:
                f.write('# 每行一个密码,以 # 或 ; 开头的行会被忽略\n')
                f.write('666\n888\n')
            self.log('未找到 passwords.txt,已创建空文件')
            self.log('请参考 passwords_example.txt 每行填一个密码')
        except Exception as e:
            self.log(f'创建密码文件失败(不影响解压): {e}')

    # ------------------------------------------------------------------
    # 日志 / 状态(全部在主线程 Tk 里处理)
    # ------------------------------------------------------------------
    def log(self, message):
        self.log_queue.put(message)

    def update_status(self):
        with self.lock:
            count = self.process_count
            active = self.active_tasks
        self.status_label.config(text=f'解压次数: {count}    待处理: {active}')

    def check_log_queue(self):
        drained = 0
        while drained < 500:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_area.insert(tk.END, msg + '\n')
            drained += 1
        if drained:
            self.log_area.see(tk.END)
        self.update_status()
        self._refresh_progress()

        if not self._close_pending and self._should_finish():
            self._close_pending = True
            self.window.after(self.close_delay_ms, self.window.destroy)
        if not self._close_pending:
            self.window.after(self.progress_interval_ms, self.check_log_queue)

    def _should_finish(self):
        if self._stay_open or not self._tasks_added or not self.seven_zip:
            return False
        with self.lock:
            if self.active_tasks > 0:
                return False
        return self.task_queue.unfinished_tasks == 0

    # ------------------------------------------------------------------
    # 任务
    # ------------------------------------------------------------------
    def add_task(self, file_path):
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            self.log(f'文件不存在: {abs_path}')
            return

        # 拖进来的是文件夹:把里面的压缩包各自拆成一个独立任务,
        # 这样才能真正用满多个工作线程,进度条也才能同时出现多条。
        # (否则整个文件夹只占 1 个任务,由一个线程从头到尾串行处理。)
        if os.path.isdir(abs_path):
            found = self._archives_in_dir(abs_path)
            if not found:
                self.log(f'文件夹里没有找到压缩包: {abs_path}')
                return
            self.log(f'文件夹 {os.path.basename(abs_path)} 里找到 '
                     f'{len(found)} 个压缩包,已拆分为并行任务')
            for item in sorted(found):
                self._enqueue(item)
            return

        self._enqueue(abs_path)

    def _enqueue(self, abs_path):
        """把一个压缩包放入任务队列。"""
        with self.lock:
            self.task_queue.put(abs_path)
            self.active_tasks += 1
            self._tasks_added = True
        self.log(f'已添加任务: {abs_path}')

    def process_worker(self):
        while True:
            current_path = self.task_queue.get()
            try:
                self.process_single_file(current_path)
            except Exception as e:
                # 兜底:任何未预料的异常都不该打死线程、丢掉后续任务
                self.log(f'处理异常({os.path.basename(current_path)}): '
                         f'{type(e).__name__}: {e}')
            finally:
                self.task_queue.task_done()
                with self.lock:
                    self.active_tasks -= 1

    # ------------------------------------------------------------------
    # 格式识别
    # ------------------------------------------------------------------
    def classify(self, path):
        """返回 '7z'/'zip'/'rar'/'polyglot' 等;不是压缩包则返回 None。

        两步走:
          1) 先看文件头 —— 快,且能识别绝大多数正常压缩包;
          2) 文件头不是压缩包时,再让 7z 扫一遍 —— 用于识别"图种"
             (JPG/PNG 里附带了 ZIP/RAR 的情况)。
             7z 对普通图片/文本会直接失败(rc=2),因此不会误伤。
        """
        if not os.path.isfile(path):
            return None
        kind = detect_format(path)
        if kind is not None:
            if kind == 'zip' and self.protect_documents and is_document_container(path):
                return None
            return kind
        # 文件头不认识:分卷的"非首卷"(.002/.r01…)本身就是续卷数据,
        # 没有标准文件头,但它是同一组压缩包的一部分,要认。
        info = volume_info(os.path.basename(path))
        if info and len(self.volume_group(path)) > 1:
            return 'volume'
        # 再考虑图种(JPG/PNG 里附带压缩包)
        if self.scan_polyglot and self._looks_like_media(path) and self._7z_can_list(path):
            return 'polyglot'
        return None

    @staticmethod
    def _looks_like_media(path):
        """只在图片/媒体类文件上做图种探测,避免对每个普通文件都起一次 7z。

        图种几乎都是"图片里塞压缩包"这一种形态,所以限定在这些扩展名上,
        既覆盖了实际场景,也避免在大目录里平白多跑几千次进程。
        """
        ext = os.path.splitext(path)[1].lower()
        return ext in POLYGLOT_EXTS

    def _7z_can_list(self, path):
        """用 7z 的列表模式判断该文件里是否含可解压内容(图种探测)。

        只读元数据,不解压,开销很小。普通图片/文本会返回非 0,故不会误判。
        """
        if not self.seven_zip:
            return False
        kwargs = {}
        flag = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        if flag:
            kwargs['creationflags'] = flag
        try:
            r = subprocess.run([self.seven_zip, 'l', '-ba', path],
                               stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL,
                               timeout=self.scan_timeout, **kwargs)
        except Exception:
            return False
        return r.returncode == 0

    def volume_group(self, main_path):
        """列出与主文件同属一组的分卷(含主文件本身)。"""
        directory = os.path.dirname(os.path.abspath(main_path))
        name = os.path.basename(main_path)
        info = volume_info(name)
        if not info:
            return [os.path.abspath(main_path)]
        base, kind, _ = info
        group = [os.path.abspath(main_path)]
        try:
            entries = os.listdir(directory)
        except OSError:
            return group
        for entry in entries:
            if entry == name:
                continue
            full = os.path.join(directory, entry)
            if not os.path.isfile(full):
                continue
            other = volume_info(entry)
            if other and other[0] == base and other[1] == kind:
                group.append(os.path.abspath(full))
        return sorted(group)

    # ------------------------------------------------------------------
    # 处理单个输入
    # ------------------------------------------------------------------
    def process_single_file(self, initial_path):
        """处理一个输入(压缩包或文件夹)。

        队列里每项带一个"嵌套深度":
          * 命令行传入的、以及文件夹里扫到的压缩包,深度都是 0;
          * 只有"解压产物里再发现的压缩包"才让深度 +1。
        max_attempts 限制的是**嵌套层数**(防止畸形包无限套娃),
        而不是"总共处理几个包"—— 否则文件夹里包一多就会被莫名截断。
        """
        file_queue = queue.Queue()
        file_queue.put((os.path.abspath(initial_path), 0))
        visited = set()
        processed = 0
        max_depth = 0

        while not file_queue.empty():
            current_path, depth = file_queue.get()
            if current_path in visited:
                continue
            visited.add(current_path)
            max_depth = max(max_depth, depth)

            if os.path.isdir(current_path):
                # 目录:扫出里面的压缩包。它们和命令行传入的包同级(深度不变)。
                for found in self._archives_in_dir(current_path):
                    file_queue.put((found, depth))
                continue

            # 用户可能直接拖来一个非首卷(.002/.r01),7z 打不开,
            # 这里自动换成同组的首卷再解。
            current_path = self._resolve_first_volume(current_path)

            kind = self.classify(current_path)
            if kind is None:
                continue

            if kind == 'rar' and not self.rar_capable:
                self.log(f'跳过 RAR(当前引擎不支持): {os.path.basename(current_path)}')
                self.log('  请把完整版 7-Zip 的 7z.exe 和 7z.dll 放到本程序同目录;')
                self.log('  原文件已保留,不会被删除。')
                continue

            if kind == 'polyglot':
                self.log(f'检测到图种(文件里含压缩包): {os.path.basename(current_path)}')
            elif kind == 'volume':
                self.log(f'检测到分卷(非首卷,将改用首卷): '
                         f'{os.path.basename(current_path)}')
            else:
                self.log(f'检测到压缩包[{kind}]: {os.path.basename(current_path)}')
            success, created, used_pwd = self.extract_archive(current_path)
            if not success:
                self.log(f'未能解压(原文件保留): {os.path.basename(current_path)}')
                continue

            processed += 1
            # 嵌套层数到上限就不再往里钻(防畸形包无限套娃)
            if depth + 1 >= self.max_attempts:
                continue
            # 递归条件(按需求):
            #   a) 解压结果里只有"一个"压缩包 -> 继续解;
            #   b) 解压结果是一组分卷(001/r01 加若干分包)-> 继续解。
            # 其他情况(混着普通文件、多个互不相关的包)就停在这里,
            # 避免把用户原本放在目录里的压缩包也一并解掉。
            for item in self._recurse_targets(created):
                file_queue.put((item, depth + 1))

        self.log(f'完成: {os.path.basename(initial_path)}'
                 f'(共解压 {processed} 个包,最深嵌套 {max_depth} 层)')

    def _recurse_targets(self, created):
        """判断本次解压结果是否该继续解,返回要入队的压缩包列表。

        只有两种情况继续:
          * 结果里恰好一个压缩包;
          * 结果是一组分卷(同一词干 + 同一分卷类型)。
        """
        archives = []
        others = 0
        for item in created:
            if os.path.isfile(item):
                if self.classify(item):
                    archives.append(item)
                else:
                    others += 1
            else:
                others += 1

        if not archives:
            return []
        # 情况 a:只有一个压缩包,且没有别的普通文件
        if len(archives) == 1 and others == 0:
            return archives
        # 情况 b:是一组分卷(且没有别的普通文件)
        if others == 0 and len(archives) > 1:
            groups = {}
            for path in archives:
                info = volume_info(os.path.basename(path))
                if not info:
                    return []          # 有非分卷成员 -> 不算"一组分卷"
                groups.setdefault((info[0], info[1]), []).append(path)
            if len(groups) == 1:
                # 只从首卷进入(7z 打不开 .002/.003)
                first = self._first_volume(next(iter(groups.values())))
                return [first] if first else []
        return []

    @staticmethod
    def _first_volume(paths):
        """从一组分卷里挑出首卷(7z 只能从首卷开始解)。"""
        best = None
        for path in paths:
            info = volume_info(os.path.basename(path))
            if not info:
                continue
            num = info[2]
            if best is None or num < best[0]:
                best = (num, path)
        return best[1] if best else None

    def _resolve_first_volume(self, path):
        """如果 path 是分卷的非首卷,返回同组首卷;否则原样返回。"""
        info = volume_info(os.path.basename(path))
        if not info:
            return path
        group = self.volume_group(path)
        if len(group) <= 1:
            return path
        first = self._first_volume(group)
        if first and os.path.normcase(first) != os.path.normcase(path):
            self.log(f'{os.path.basename(path)} 是分卷的非首卷,'
                     f'改用首卷 {os.path.basename(first)} 解压')
        return first or path

    def _archives_in_dir(self, directory):
        """递归找出目录里的压缩包。"""
        found = []
        for root, dirs, files in os.walk(directory):
            for name in files:
                full = os.path.join(root, name)
                if self.classify(full):
                    found.append(full)
        return found

    # ------------------------------------------------------------------
    # 解压
    # ------------------------------------------------------------------
    def extract_archive(self, archive_path):
        """解压一个压缩包。返回 (是否成功, 落到目标目录的路径集合, 用到的密码)。"""
        archive_path = os.path.abspath(archive_path)
        target_dir = os.path.dirname(archive_path)

        tmp_dir, same_volume = self._make_temp_dir(target_dir)
        if tmp_dir is None:
            self.log('无法创建临时目录,跳过')
            return False, set(), None

        base = os.path.basename(archive_path)
        task_id = self._begin_progress(base)
        try:
            used_pwd, ok = self._try_passwords(archive_path, tmp_dir, base, task_id)
            if not ok:
                return False, set(), None

            created = set()
            failed = []
            # 7z 报的 0-100% 只覆盖"解压"这一段,后面还要把文件搬出临时目录。
            # 所以进度条切成两段:0-80% 解压、80-100% 搬运,
            # 免得文件多时进度条在搬运阶段一动不动。
            items = sorted(os.listdir(tmp_dir))
            total_items = max(1, len(items))
            for index, name in enumerate(items):
                self._merge_move(os.path.join(tmp_dir, name),
                                 target_dir, created, failed)
                moving_pct = 80 + int((index + 1) * 20 / total_items)
                self._set_progress(moving_pct, f'{base} [搬移文件]', task_id)

            if failed:
                # 关键:有文件没搬出来 -> 保留原包,绝不删
                self.log(f'有 {len(failed)} 项没能移出临时目录,原压缩包保留不删')
                for src, err in failed[:5]:
                    self.log(f'  移动失败: {os.path.basename(src)} -> {err}')
                return False, created, used_pwd

            # 全部搬出来了才算成功 -> 删原包和分卷
            self._delete_archive(archive_path)
            with self.lock:
                self.process_count += 1
            # 控制台/日志:XXXX 解压完毕
            self.log(f'{base} 解压完毕')
            return True, created, used_pwd
        except Exception as e:
            self.log(f'解压出错: {type(e).__name__}: {e}')
            return False, set(), None
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            self._end_progress(task_id)

    def _make_temp_dir(self, target_dir):
        """在目标目录(同盘)建临时目录;不行就退回系统临时目录。

        注意:刻意不用 tempfile.mkdtemp —— 它生成的目录在部分受限环境下
        不继承调用方的写入权限,会导致解压直接失败。这里自己造随机名并查重。
        """
        for directory in (target_dir, tempfile.gettempdir()):
            for _ in range(20):
                name = self.temp_dir_prefix + uuid.uuid4().hex[:10]
                path = os.path.join(directory, name)
                try:
                    os.makedirs(path)  # exist_ok=False:撞名就换一个
                    return path, directory == target_dir
                except FileExistsError:
                    continue
                except OSError:
                    break
        return None, False

    def _try_passwords(self, archive_path, tmp_dir, base='', task_id=None):
        """依次尝试无密码和密码表里的密码。返回 (用到的密码, 是否成功)。"""
        name = base or os.path.basename(archive_path)
        for pwd in [None] + list(self.passwords):
            self._empty_dir(tmp_dir)
            label = f'{name} [密码 {pwd}]' if pwd else f'{name} [无密码]'
            self.log(f'尝试{"密码 " + pwd if pwd else "无密码"}解压')
            try:
                returncode, _ = self.run_7z_command(archive_path, tmp_dir, pwd,
                                                    label, task_id)
            except Exception as e:
                self.log(f'调用 7z 失败: {e}')
                return None, False
            if returncode == 0:
                return pwd, True
            # 这个密码不行,重试下一个时把进度条归零
            self._set_progress(0, label, task_id)
        self.log('所有密码都试过了,不行')
        return None, False

    @staticmethod
    def _empty_dir(directory):
        for name in os.listdir(directory):
            full = os.path.join(directory, name)
            try:
                if os.path.isdir(full) and not os.path.islink(full):
                    shutil.rmtree(full, ignore_errors=True)
                else:
                    os.remove(full)
            except OSError:
                pass

    def _merge_move(self, src, dest_dir, created, failed):
        """把 src 搬到 dest_dir:同名文件自动改名,同名目录合并。绝不覆盖。

        注意:多个工作线程可能同时往同一个目标目录搬文件,而"挑选一个没被
        占用的名字"和"真正搬过去"这两步之间必须原子,否则两个线程会算出
        同一个目标名,后搬的把先搬的覆盖掉 —— 直接丢文件。
        所以这里用 self.move_lock 把两步圈起来。
        """
        target = os.path.join(dest_dir, os.path.basename(src))
        if os.path.isdir(src):
            with self.move_lock:
                if os.path.exists(target) and not os.path.isdir(target):
                    target = self._unique_name(target)
                try:
                    os.makedirs(target, exist_ok=True)
                except OSError as e:
                    failed.append((src, e))
                    return
            created.add(target)
            try:
                children = sorted(os.listdir(src))
            except OSError:
                return
            for child in children:
                self._merge_move(os.path.join(src, child), target, created, failed)
        else:
            with self.move_lock:
                final = self._unique_name(target)
                try:
                    shutil.move(src, final)
                except Exception as e:
                    failed.append((src, e))
                    return
            created.add(final)

    @staticmethod
    def _unique_name(path):
        """若 path 已被占用,生成 xxx_1 / xxx_2 ...

        调用方必须持有 self.move_lock:并发的"选名字 -> 搬过去"不是原子的,
        两个线程会选到同一个名字并互相覆盖。
        """
        if not os.path.exists(path):
            return path
        stem, ext = os.path.splitext(path)
        index = 1
        while True:
            candidate = f'{stem}_{index}{ext}'
            if not os.path.exists(candidate):
                return candidate
            index += 1

    def _delete_archive(self, archive_path):
        """删掉原包及其分卷(只有确认解压成功才会走到这里)。"""
        for victim in self.volume_group(archive_path):
            if not os.path.exists(victim):
                continue
            try:
                os.remove(victim)
                self.log(f'删除原文件: {os.path.basename(victim)}')
            except Exception as e:
                self.log(f'删除失败({os.path.basename(victim)}): {e}')

    def run_7z_command(self, archive_path, out_dir, password=None, label='',
                       task_id=None):
        """执行 7z 解压,并实时解析进度。

        7z 用 '\\r' 原地刷新进度行(和命令行进度条一样),这里把它解析成
        百分比推给界面。用 -bsp1 让进度走 stdout;不用 read()(它会缓冲到
        填满才返回,导致进度看着像卡住),改用无缓冲的 os.read()。
        返回 (returncode, 输出文本)。
        """
        cmd = [self.seven_zip, 'x', '-y', '-aoa', '-bsp1', f'-o{out_dir}']
        if password:
            cmd.append(f'-p{password}')
        cmd.append(archive_path)

        kwargs = {}
        flag = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        if flag:
            kwargs['creationflags'] = flag

        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT, bufsize=0, **kwargs)
        chunks = []
        deadline = time.time() + self.extract_timeout
        last_progress = time.time()
        last_pct = -1
        pct_re = re.compile(rb'(\d{1,3})%')

        try:
            fd = proc.stdout.fileno()
            while True:
                try:
                    chunk = os.read(fd, 4096)
                except OSError:
                    break
                if not chunk:
                    break
                chunks.append(chunk)

                # 解析本批数据里的百分比(取最后一个,即最新进度)
                found = None
                for m in pct_re.finditer(chunk):
                    found = int(m.group(1))
                if found is not None and found != last_pct:
                    last_pct = found
                    last_progress = time.time()
                    # 7z 的 0-100% 只覆盖"解压"这一段;后面还要把文件从临时
                    # 目录搬出去。因此把解压映射到整条进度的 0-80%。
                    self._set_progress(int(found * 0.8), label, task_id)
                elif found is not None:
                    last_progress = time.time()

                now = time.time()
                if now > deadline:
                    proc.kill()
                    self.log(f'解压超时(超过 {self.extract_timeout} 秒),已中止')
                    break
                # 长时间没有任何进度输出:可能是坏包卡住
                if now - last_progress > self.progress_stall_s:
                    proc.kill()
                    self.log(f'超过 {self.progress_stall_s} 秒没有进度,已中止')
                    break
        finally:
            try:
                proc.stdout.close()
            except Exception:
                pass
            proc.wait()

        text = b''.join(chunks).decode('utf-8', 'replace')
        if proc.returncode != 0:
            # 失败(如密码不对),进度归零,交给下一次尝试
            self._set_progress(0, label, task_id)
        return proc.returncode, text

    # ------------------------------------------------------------------
    # 进度上报(worker 写 -> 主线程读并刷新界面)
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # 进度上报(worker 写 -> 主线程读并刷新界面)
    #
    # 每个正在解压的压缩包占一条进度:文件名 | 进度条 | 百分比。
    # 解压完这条就消失;多个包同时解压就并排显示多条。
    # ------------------------------------------------------------------
    def _begin_progress(self, label=''):
        """一个压缩包开始解压,注册一条独立进度。返回 task_id。"""
        with self.progress_lock:
            self._progress_seq += 1
            task_id = self._progress_seq
            self._progress[task_id] = {
                'label': label, 'percent': 0, 'updated': time.time(),
            }
        return task_id

    def _set_progress(self, percent, label=None, task_id=None):
        with self.progress_lock:
            item = self._progress.get(task_id)
            if item is not None:
                item['percent'] = max(0, min(100, int(percent)))
                if label is not None:
                    item['label'] = label
                item['updated'] = time.time()

    def _end_progress(self, task_id=None):
        """一个压缩包解压结束(成功或失败),它的那条进度随即消失。"""
        with self.progress_lock:
            self._progress.pop(task_id, None)

    def _refresh_progress(self):
        """在主线程里刷新进度区:为每个任务建/更新/删除一行。"""
        if not self.show_progress:
            return
        with self.progress_lock:
            snapshot = {k: dict(v) for k, v in self._progress.items()}
        try:
            # 已结束的任务:把对应那一行移除(进度条消失)
            for task_id in list(self.progress_rows):
                if task_id not in snapshot:
                    row = self.progress_rows.pop(task_id)
                    row['frame'].destroy()

            # 进行中的任务:建行或更新
            for task_id, info in snapshot.items():
                row = self.progress_rows.get(task_id)
                if row is None:
                    row = self._make_progress_row(task_id, info['label'])
                pct = info['percent']
                row['bar']['value'] = pct
                if pct < 80 and (time.time() - info['updated']) > 5:
                    row['pct'].config(text=f'{pct}% 等待')
                else:
                    row['pct'].config(text=f'{pct}%')
                row['name'].config(text=self._shorten(info['label']))
        except tk.TclError:
            pass

    def _make_progress_row(self, task_id, label):
        """建一行:左边文件名,中间进度条,右边百分比。

        注意:tk.Label 的 width 单位是"字符数",中英文混排时极不准确
        (中文一个字符约占 2 个英文字符宽)。所以文件名栏用一个固定像素
        宽度的容器 + 内部 Label 靠左,并用 pack_propagate(False) 锁住宽度。
        """
        frame = tk.Frame(self.progress_frame)
        frame.pack(fill=tk.X, padx=3, pady=1)

        holder = tk.Frame(frame, width=PROGRESS_NAME_PX, height=20)
        holder.pack(side=tk.LEFT)
        holder.pack_propagate(False)          # 锁定像素宽度,不被内容撑开
        name = tk.Label(holder, text=self._shorten(label), anchor='w')
        name.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 百分比放右边,先 pack 保证它不会被进度条挤掉
        pct = tk.Label(frame, text='0%', width=6, anchor='e')
        pct.pack(side=tk.RIGHT)

        bar = ttk.Progressbar(frame, orient='horizontal', mode='determinate',
                              maximum=100)
        bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 4))

        row = {'frame': frame, 'name': name, 'bar': bar, 'pct': pct}
        self.progress_rows[task_id] = row
        return row

    @staticmethod
    def _shorten(text, limit=PROGRESS_NAME_CHARS):
        """名字太长就中间省略,保证能看清开头和结尾。"""
        if len(text) <= limit:
            return text
        head = limit // 2
        tail = limit - head - 1
        return text[:head] + '…' + text[-tail:]


if __name__ == '__main__':
    app = AutoUnzipApp()
    app.window.mainloop()

