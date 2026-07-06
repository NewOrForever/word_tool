#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import queue
import sys
import threading
import zipfile
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
from tkinter.ttk import Progressbar, Style

from remove_word_headers_footers import collect_targets, process_docx


APP_TITLE = "Word 页眉页脚清理工具"
MODE_LABELS = {
    "header": "只去页眉",
    "footer": "只去页脚",
    "both": "页眉页脚都去",
}


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
    # 工作线程消息队列
    message_queue: queue.Queue

    # 初始化界面和状态
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.target_var = StringVar()
        self.mode_var = StringVar(value="both")
        self.recursive_var = BooleanVar(value=True)
        self.target_type_var = StringVar(value="folder")
        self.message_queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.root.title(APP_TITLE)
        self.root.geometry("760x560")
        self.root.minsize(720, 520)
        self.root.resizable(True, True)

        self._configure_style()
        self._build_ui()
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

    # 创建处理选项区域
    def _build_options_panel(self, parent: Frame) -> None:
        panel = LabelFrame(parent, text="处理选项", padx=12, pady=10)
        panel.pack(fill="x", pady=(0, 12))

        mode_row = Frame(panel)
        mode_row.pack(fill="x")

        for mode, label in MODE_LABELS.items():
            Radiobutton(mode_row, text=label, value=mode, variable=self.mode_var).pack(side="left", padx=(0, 18))

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
            self.target_var.set(selected)

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

        return target.resolve()

    # 开始后台处理
    def _start_processing(self) -> None:
        target = self._validate_target()
        if target is None:
            return

        mode = self.mode_var.get()
        recursive = bool(self.recursive_var.get()) and self.target_type_var.get() == "folder"

        self.start_button.config(state=DISABLED)
        self.progress["value"] = 0
        self.status_label.config(text="处理中")
        self.log_text.delete("1.0", END)
        self._append_log(f"处理模式：{MODE_LABELS[mode]}")
        self._append_log(f"处理路径：{target}")
        self._append_log("提示：处理完成后会直接覆盖原文件")

        # 后台线程执行文件处理，避免界面卡住
        self.worker_thread = threading.Thread(
            target=self._run_processing,
            args=(target, mode, recursive),
            daemon=True,
        )
        self.worker_thread.start()

    # 后台执行 Word 处理
    def _run_processing(self, target: Path, mode: str, recursive: bool) -> None:
        try:
            targets = collect_targets(target, recursive)
            if not targets:
                self.message_queue.put(("error", "没有找到可处理的新版 Word 文件"))
                return

            success_count = 0
            skip_count = 0
            total = len(targets)

            for index, source in enumerate(targets, start=1):
                result = process_docx(source, mode, None, True)
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
    }

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content.encode("utf-8"))


# 程序入口
def main() -> int:
    if "--self-test" in sys.argv:
        return run_self_test()

    root = Tk()
    HeaderFooterCleanerApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
