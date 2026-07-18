#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import queue
import ctypes
import importlib
import platform
import sys
import threading
import zipfile
from base64 import b64decode
from pathlib import Path
from tkinter import (
    BooleanVar,
    Button,
    DISABLED,
    END,
    Entry,
    Frame,
    Label,
    LabelFrame,
    NORMAL,
    Radiobutton,
    StringVar,
    Tk,
    filedialog,
    messagebox,
)
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import Combobox, Progressbar, Style
from xml.etree import ElementTree as ET

from remove_word_headers_footers import WORD_NS, collect_targets, process_docx

DND_FILES = None
TkinterDnD = None


# 加载拖拽组件，打包后也从临时资源目录查找
def load_tkinterdnd() -> bool:
    global DND_FILES, TkinterDnD
    if DND_FILES is not None and TkinterDnD is not None:
        return True

    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir and bundle_dir not in sys.path:
        sys.path.insert(0, bundle_dir)

    try:
        module = importlib.import_module("tkinterdnd2")
    except Exception:
        return False

    DND_FILES = module.DND_FILES
    TkinterDnD = module.TkinterDnD
    return True


# 创建主窗口，拖拽窗口创建失败时退回普通窗口
def create_root_window() -> Tk:
    if not load_tkinterdnd():
        return Tk()

    try:
        return TkinterDnD.Tk()
    except Exception:
        return Tk()


APP_TITLE = "Word 页眉页脚清理工具"
MODE_LABELS = {
    "header": "只去页眉",
    "footer": "只去页脚",
    "both": "页眉页脚都去",
    "set_header": "修改/增加页眉",
    "set_footer": "修改/增加页脚",
    "set_both": "页眉页脚都修改",
}
# 界面位置文字和文档对齐值的对应关系
ALIGNMENT_LABELS = {"左边": "left", "居中": "center", "右边": "right"}


# 把界面位置文字转换成文档对齐值
def alignment_value(label: str) -> str:
    try:
        return ALIGNMENT_LABELS[label]
    except KeyError as exc:
        raise ValueError(f"不支持的位置：{label}") from exc


class HeaderFooterCleanerApp:
    # 主窗口对象
    root: Tk
    # 输入路径变量
    target_var: StringVar
    # 处理模式变量
    mode_var: StringVar
    # 是否递归处理变量
    recursive_var: BooleanVar
    # 当前处理类型变量
    target_type_var: StringVar
    # 页眉内容变量
    header_text_var: StringVar
    # 页脚内容变量
    footer_text_var: StringVar
    # 页眉位置变量
    header_alignment_var: StringVar
    # 页脚位置变量
    footer_alignment_var: StringVar
    # 工作线程消息队列
    message_queue: queue.Queue
    # 拖拽回调保存列表
    drop_callbacks: list[object]

    # 初始化界面和状态
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.target_var = StringVar()
        self.mode_var = StringVar(value="both")
        self.recursive_var = BooleanVar(value=True)
        self.target_type_var = StringVar(value="folder")
        self.header_text_var = StringVar()
        self.footer_text_var = StringVar()
        self.header_alignment_var = StringVar(value="左边")
        self.footer_alignment_var = StringVar(value="右边")
        self.message_queue = queue.Queue()
        self.drop_callbacks = []
        self.worker_thread: threading.Thread | None = None

        self.root.title(APP_TITLE)
        self.root.geometry("760x560")
        self.root.minsize(720, 520)
        self.root.resizable(True, True)

        self._configure_style()
        self._build_ui()
        self._enable_drop_upload()
        self._poll_messages()

    # 设置窗口控件样式
    def _configure_style(self) -> None:
        style = Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Horizontal.TProgressbar", thickness=14)

    # 创建主界面
    def _build_ui(self) -> None:
        root_frame = Frame(self.root, padx=18, pady=16)
        root_frame.pack(fill="both", expand=True)

        title = Label(root_frame, text=APP_TITLE, font=("Microsoft YaHei UI", 18, "bold"))
        title.pack(anchor="w")

        subtitle = Label(
            root_frame,
            text="选择 Word 文件或文件夹，处理完成后会直接覆盖原文件",
            font=("Microsoft YaHei UI", 10),
            fg="#555555",
        )
        subtitle.pack(anchor="w", pady=(4, 14))

        self._build_target_panel(root_frame)
        self._build_options_panel(root_frame)
        self._build_action_panel(root_frame)
        self._build_log_panel(root_frame)

    # 创建路径选择区域
    def _build_target_panel(self, parent: Frame) -> None:
        panel = LabelFrame(parent, text="处理范围", padx=12, pady=10)
        panel.pack(fill="x", pady=(0, 12))

        type_row = Frame(panel)
        type_row.pack(fill="x", pady=(0, 8))

        Radiobutton(
            type_row,
            text="文件夹",
            value="folder",
            variable=self.target_type_var,
            command=self._clear_target,
        ).pack(side="left")
        Radiobutton(
            type_row,
            text="单个 Word 文件",
            value="file",
            variable=self.target_type_var,
            command=self._clear_target,
        ).pack(side="left", padx=(16, 0))

        path_row = Frame(panel)
        path_row.pack(fill="x")

        self.path_entry = Entry(path_row, textvariable=self.target_var)
        self.path_entry.pack(side="left", fill="x", expand=True)

        Button(path_row, text="选择", width=10, command=self._choose_target).pack(side="left", padx=(8, 0))

        drop_hint = Label(panel, text="也可以把 Word 文件或文件夹拖到窗口里", fg="#777777")
        drop_hint.pack(anchor="w", pady=(8, 0))

    # 创建处理选项区域
    def _build_options_panel(self, parent: Frame) -> None:
        panel = LabelFrame(parent, text="处理选项", padx=12, pady=10)
        panel.pack(fill="x", pady=(0, 12))

        mode_row = Frame(panel)
        mode_row.pack(fill="x")

        for mode, label in MODE_LABELS.items():
            Radiobutton(mode_row, text=label, value=mode, variable=self.mode_var).pack(side="left", padx=(0, 18))

        header_row = Frame(panel)
        header_row.pack(fill="x", pady=(10, 0))

        Label(header_row, text="页眉内容", width=8, anchor="w").pack(side="left")
        self.header_entry = Entry(header_row, textvariable=self.header_text_var)
        self.header_entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        Label(header_row, text="位置", width=5, anchor="e").pack(side="left", padx=(10, 4))
        self.header_alignment_box = Combobox(
            header_row,
            textvariable=self.header_alignment_var,
            values=tuple(ALIGNMENT_LABELS),
            state="readonly",
            width=6,
        )
        self.header_alignment_box.pack(side="left")

        footer_row = Frame(panel)
        footer_row.pack(fill="x", pady=(8, 0))

        Label(footer_row, text="页脚内容", width=8, anchor="w").pack(side="left")
        self.footer_entry = Entry(footer_row, textvariable=self.footer_text_var)
        self.footer_entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        Label(footer_row, text="位置", width=5, anchor="e").pack(side="left", padx=(10, 4))
        self.footer_alignment_box = Combobox(
            footer_row,
            textvariable=self.footer_alignment_var,
            values=tuple(ALIGNMENT_LABELS),
            state="readonly",
            width=6,
        )
        self.footer_alignment_box.pack(side="left")

        recursive_row = Frame(panel)
        recursive_row.pack(fill="x", pady=(10, 0))

        Radiobutton(
            recursive_row,
            text="处理子文件夹",
            value=True,
            variable=self.recursive_var,
        ).pack(side="left")
        Radiobutton(
            recursive_row,
            text="只处理当前文件夹",
            value=False,
            variable=self.recursive_var,
        ).pack(side="left", padx=(18, 0))

    # 创建操作按钮和进度条
    def _build_action_panel(self, parent: Frame) -> None:
        panel = Frame(parent)
        panel.pack(fill="x", pady=(0, 12))

        self.start_button = Button(panel, text="开始处理", width=14, command=self._start_processing)
        self.start_button.pack(side="left")

        self.progress = Progressbar(panel, mode="determinate", maximum=100)
        self.progress.pack(side="left", fill="x", expand=True, padx=(14, 0))

        self.status_label = Label(panel, text="等待开始", width=18, anchor="e", fg="#555555")
        self.status_label.pack(side="left", padx=(12, 0))

    # 创建日志区域
    def _build_log_panel(self, parent: Frame) -> None:
        panel = LabelFrame(parent, text="处理结果", padx=10, pady=8)
        panel.pack(fill="both", expand=True)

        self.log_text = ScrolledText(panel, height=14, wrap="word", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
        self._append_log("请选择文件或文件夹后点击“开始处理”")

    # 清空已选择路径
    def _clear_target(self) -> None:
        self.target_var.set("")

    # 打开文件或文件夹选择框
    def _choose_target(self) -> None:
        if self.target_type_var.get() == "file":
            selected = filedialog.askopenfilename(
                title="选择 Word 文件",
                filetypes=[
                    ("Word 文件", "*.docx *.docm *.dotx *.dotm"),
                    ("所有文件", "*.*"),
                ],
            )
        else:
            selected = filedialog.askdirectory(title="选择 Word 文件夹")

        if selected:
            self._set_target_from_path(Path(selected))

    # 启用 Windows 文件拖拽上传
    def _enable_drop_upload(self) -> None:
        if self._enable_tkinterdnd_upload():
            return
        self._enable_windows_drop_upload()

    # 启用 TkDND 文件拖拽上传
    def _enable_tkinterdnd_upload(self) -> bool:
        if not load_tkinterdnd():
            return False

        registered = False
        for widget in self._iter_widgets(self.root):
            if not hasattr(widget, "drop_target_register"):
                continue
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._handle_drop_event)
                registered = True
            except Exception:
                continue
        return registered

    # 处理 TkDND 拖拽事件
    def _handle_drop_event(self, event) -> None:
        self._handle_dropped_paths(self._parse_dropped_paths(event.data))

    # 解析拖拽事件中的路径列表
    def _parse_dropped_paths(self, data: str) -> list[str]:
        try:
            return [path for path in self.root.tk.splitlist(data) if path]
        except Exception:
            return [data] if data else []

    # 启用 Windows 原生文件拖拽上传
    def _enable_windows_drop_upload(self) -> None:
        if sys.platform != "win32":
            return

        try:
            import ctypes
            from ctypes import wintypes
        except Exception:
            return

        WM_DROPFILES = 0x0233
        GWLP_WNDPROC = -4
        LRESULT = ctypes.c_ssize_t
        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32
        set_window_long = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)
        call_window_proc = user32.CallWindowProcW
        set_window_long.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        set_window_long.restype = ctypes.c_void_p
        call_window_proc.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        call_window_proc.restype = LRESULT

        # 所有主要控件都接收拖拽，避免拖到输入框或日志区时失效
        self.root.update_idletasks()
        seen_hwnds: set[int] = set()
        for widget in self._iter_widgets(self.root):
            try:
                hwnd = widget.winfo_id()
                if hwnd in seen_hwnds:
                    continue
                seen_hwnds.add(hwnd)
                shell32.DragAcceptFiles(hwnd, True)
            except Exception:
                continue

            old_proc_holder: list[ctypes.c_void_p] = [ctypes.c_void_p()]

            def drop_proc(hwnd_param, message, w_param, l_param, old_proc_holder=old_proc_holder):
                if message == WM_DROPFILES:
                    paths = self._read_drop_paths(shell32, w_param)
                    shell32.DragFinish(w_param)
                    self.root.after(0, lambda: self._handle_dropped_paths(paths))
                    return 0
                return call_window_proc(old_proc_holder[0], hwnd_param, message, w_param, l_param)

            callback = WNDPROC(drop_proc)
            old_proc_holder[0] = set_window_long(hwnd, GWLP_WNDPROC, ctypes.cast(callback, ctypes.c_void_p))
            self.drop_callbacks.append(callback)

    # 遍历窗口中的全部控件
    def _iter_widgets(self, widget) -> list:
        widgets = [widget]
        for child in widget.winfo_children():
            widgets.extend(self._iter_widgets(child))
        return widgets

    # 读取系统拖拽消息里的路径列表
    def _read_drop_paths(self, shell32, drop_handle) -> list[str]:
        count = shell32.DragQueryFileW(drop_handle, 0xFFFFFFFF, None, 0)
        paths: list[str] = []
        for index in range(count):
            length = shell32.DragQueryFileW(drop_handle, index, None, 0) + 1
            buffer = ctypes.create_unicode_buffer(length)
            shell32.DragQueryFileW(drop_handle, index, buffer, length)
            paths.append(buffer.value)
        return paths

    # 处理拖拽进来的文件或文件夹路径
    def _handle_dropped_paths(self, paths: list[str]) -> None:
        if not paths:
            return

        self._set_target_from_path(Path(paths[0]))
        if len(paths) > 1:
            self._append_log("已选择第一个拖入的文件或文件夹")

    # 根据路径自动切换文件或文件夹模式
    def _set_target_from_path(self, target: Path) -> None:
        if target.is_dir():
            self.target_type_var.set("folder")
            self.target_var.set(str(target))
            return
        if target.is_file():
            self.target_type_var.set("file")
            self.target_var.set(str(target))
            return
        messagebox.showwarning("路径不正确", "拖入的路径不是有效的文件或文件夹")

    # 校验用户输入
    def _validate_target(self) -> Path | None:
        target_text = self.target_var.get().strip()
        if not target_text:
            messagebox.showwarning("缺少路径", "请先选择 Word 文件或文件夹")
            return None

        target = Path(target_text)
        if self.target_type_var.get() == "file" and not target.is_file():
            messagebox.showwarning("路径不正确", "请选择一个有效的 Word 文件")
            return None
        if self.target_type_var.get() == "folder" and not target.is_dir():
            messagebox.showwarning("路径不正确", "请选择一个有效的文件夹")
            return None
        mode = self.mode_var.get()
        if mode in {"set_header", "set_both"} and not self.header_text_var.get().strip():
            messagebox.showwarning("缺少页眉内容", "请填写要写入的页眉内容")
            return None
        if mode in {"set_footer", "set_both"} and not self.footer_text_var.get().strip():
            messagebox.showwarning("缺少页脚内容", "请填写要写入的页脚内容")
            return None

        return target.resolve()

    # 开始后台处理
    def _start_processing(self) -> None:
        target = self._validate_target()
        if target is None:
            return

        mode = self.mode_var.get()
        header_text = self.header_text_var.get().strip() if mode in {"set_header", "set_both"} else None
        footer_text = self.footer_text_var.get().strip() if mode in {"set_footer", "set_both"} else None
        header_alignment = alignment_value(self.header_alignment_var.get())
        footer_alignment = alignment_value(self.footer_alignment_var.get())
        recursive = bool(self.recursive_var.get()) and self.target_type_var.get() == "folder"

        self.start_button.config(state=DISABLED)
        self.progress["value"] = 0
        self.status_label.config(text="处理中")
        self.log_text.delete("1.0", END)
        self._append_log(f"处理模式：{MODE_LABELS[mode]}")
        self._append_log(f"处理路径：{target}")
        if header_text is not None:
            self._append_log(f"页眉内容：{header_text}")
            self._append_log(f"页眉位置：{self.header_alignment_var.get()}")
        if footer_text is not None:
            self._append_log(f"页脚内容：{footer_text}")
            self._append_log(f"页脚位置：{self.footer_alignment_var.get()}")
        self._append_log("提示：处理完成后会直接覆盖原文件")

        # 后台线程执行文件处理，避免界面卡住
        self.worker_thread = threading.Thread(
            target=self._run_processing,
            args=(target, mode, recursive, header_text, footer_text, header_alignment, footer_alignment),
            daemon=True,
        )
        self.worker_thread.start()

    # 后台执行 Word 处理
    def _run_processing(
        self,
        target: Path,
        mode: str,
        recursive: bool,
        header_text: str | None,
        footer_text: str | None,
        header_alignment: str,
        footer_alignment: str,
    ) -> None:
        try:
            targets = collect_targets(target, recursive)
            if not targets:
                self.message_queue.put(("error", "没有找到可处理的新版 Word 文件"))
                return

            success_count = 0
            skip_count = 0
            total = len(targets)

            for index, source in enumerate(targets, start=1):
                result = process_docx(
                    source,
                    mode,
                    None,
                    True,
                    header_text=header_text,
                    footer_text=footer_text,
                    header_alignment=header_alignment,
                    footer_alignment=footer_alignment,
                )
                if result.changed:
                    success_count += 1
                    self.message_queue.put(("log", f"[成功] {source}"))
                else:
                    skip_count += 1
                    self.message_queue.put(("log", f"[跳过] {source}：{result.message}"))

                progress_value = int(index / total * 100)
                self.message_queue.put(("progress", progress_value))

            summary = f"完成：成功 {success_count} 个，跳过 {skip_count} 个"
            self.message_queue.put(("done", summary))
        except Exception as exc:
            self.message_queue.put(("error", f"处理失败：{exc}"))

    # 轮询后台线程消息
    def _poll_messages(self) -> None:
        try:
            while True:
                kind, payload = self.message_queue.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "progress":
                    self.progress["value"] = payload
                elif kind == "done":
                    self.progress["value"] = 100
                    self.status_label.config(text="已完成")
                    self.start_button.config(state=NORMAL)
                    self._append_log(payload)
                    messagebox.showinfo("处理完成", payload)
                elif kind == "error":
                    self.status_label.config(text="处理失败")
                    self.start_button.config(state=NORMAL)
                    self._append_log(payload)
                    messagebox.showerror("处理失败", payload)
        except queue.Empty:
            pass

        self.root.after(100, self._poll_messages)

    # 向日志区域追加一行文本
    def _append_log(self, text: str) -> None:
        self.log_text.insert(END, text + "\n")
        self.log_text.see(END)


# 执行无界面的自检逻辑
def run_self_test() -> int:
    if not load_tkinterdnd():
        print("SELF TEST FAIL: 拖拽组件未加载")
        return 1

    root = create_root_window()
    root.withdraw()
    root.destroy()

    sample_dir = Path("work/gui_self_test")
    source = sample_dir / "sample.docx"
    sample_dir.mkdir(parents=True, exist_ok=True)
    create_sample_docx(source)

    result = process_docx(source, "footer", None, True)
    if not result.changed:
        print("SELF TEST FAIL: 未处理文件")
        return 1

    with zipfile.ZipFile(source) as zf:
        names = set(zf.namelist())
        combined = "\n".join(
            zf.read(name).decode("utf-8", errors="ignore")
            for name in zf.namelist()
            if name.endswith(".xml") or name.endswith(".rels")
        ) + "\n" + "\n".join(names)

    if "footerReference" in combined or "word/footer1.xml" in combined:
        print("SELF TEST FAIL: 页脚仍然存在")
        return 1
    if "headerReference" not in combined or "word/header1.xml" not in combined:
        print("SELF TEST FAIL: 页眉被误删")
        return 1

    update_source = sample_dir / "update_sample.docx"
    create_sample_docx(update_source)
    update_result = process_docx(
        update_source,
        "set_both",
        None,
        True,
        header_text="测试修改页眉",
        footer_text="测试修改页脚",
        header_alignment="center",
        footer_alignment="left",
    )
    if not update_result.changed:
        print("SELF TEST FAIL: 未修改页眉页脚")
        return 1

    with zipfile.ZipFile(update_source) as zf:
        update_names = set(zf.namelist())
        header_root = ET.fromstring(zf.read("word/header1.xml"))
        footer_root = ET.fromstring(zf.read("word/footer1.xml"))
        footer_text = "".join(text.text or "" for text in footer_root.iter(f"{{{WORD_NS}}}t"))
        header_alignment = header_root.find(f".//{{{WORD_NS}}}p/{{{WORD_NS}}}pPr/{{{WORD_NS}}}jc")
        footer_alignment = footer_root.find(f".//{{{WORD_NS}}}p/{{{WORD_NS}}}pPr/{{{WORD_NS}}}jc")

    if footer_text != "测试修改页脚":
        print("SELF TEST FAIL: 页脚内容未修改")
        return 1
    if header_alignment is None or header_alignment.attrib.get(f"{{{WORD_NS}}}val") != "center":
        print("SELF TEST FAIL: 页眉未居中")
        return 1
    if footer_alignment is None or footer_alignment.attrib.get(f"{{{WORD_NS}}}val") != "left":
        print("SELF TEST FAIL: 页脚未左对齐")
        return 1
    if "word/_rels/footer1.xml.rels" in update_names:
        print("SELF TEST FAIL: 旧页脚关联未清理")
        return 1
    if "word/media/old-footer.png" in update_names:
        print("SELF TEST FAIL: 旧页脚图片未清理")
        return 1

    print("SELF TEST PASS")
    return 0


# 创建用于打包后验证的示例 Word 文件
def create_sample_docx(path: Path) -> None:
    files = {
        "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""",
        "word/document.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>正文内容</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader"/>
      <w:footerReference w:type="default" r:id="rIdFooter"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>""",
        "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>""",
        "word/header1.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>测试页眉</w:t></w:r></w:p></w:hdr>""",
        "word/footer1.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>测试页脚</w:t></w:r></w:p></w:ftr>""",
        "word/_rels/footer1.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="file:///D:/old/image.png" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/old-footer.png"/>
</Relationships>""",
        "word/media/old-footer.png": b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        ),
    }

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            data = content if isinstance(content, bytes) else content.encode("utf-8")
            zf.writestr(name, data)


# 程序入口
def main() -> int:
    if "--self-test" in sys.argv:
        return run_self_test()

    root = create_root_window()
    HeaderFooterCleanerApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
