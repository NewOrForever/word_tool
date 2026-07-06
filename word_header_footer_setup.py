#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from tkinter import BooleanVar, Button, Checkbutton, Entry, Frame, Label, StringVar, Tk, filedialog, messagebox


APP_NAME = "WordHeaderFooterCleaner"
APP_TITLE = "Word 页眉页脚清理工具安装包"
EXE_NAME = "WordHeaderFooterCleaner.exe"


class InstallerApp:
    # 主窗口对象
    root: Tk
    # 安装目录变量
    install_dir_var: StringVar
    # 是否创建桌面快捷方式变量
    shortcut_var: BooleanVar

    # 初始化安装界面
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.install_dir_var = StringVar(value=str(self._default_install_dir()))
        self.shortcut_var = BooleanVar(value=True)

        self.root.title(APP_TITLE)
        self.root.geometry("560x220")
        self.root.resizable(False, False)

        self._build_ui()

    # 创建安装界面
    def _build_ui(self) -> None:
        root_frame = Frame(self.root, padx=18, pady=16)
        root_frame.pack(fill="both", expand=True)

        Label(root_frame, text=APP_TITLE, font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w")
        Label(root_frame, text="选择安装位置后点击安装", fg="#555555").pack(anchor="w", pady=(4, 16))

        path_row = Frame(root_frame)
        path_row.pack(fill="x")

        Entry(path_row, textvariable=self.install_dir_var).pack(side="left", fill="x", expand=True)
        Button(path_row, text="浏览", width=8, command=self._choose_dir).pack(side="left", padx=(8, 0))

        Checkbutton(root_frame, text="创建桌面快捷方式", variable=self.shortcut_var).pack(anchor="w", pady=(14, 18))

        action_row = Frame(root_frame)
        action_row.pack(fill="x")
        Button(action_row, text="安装", width=12, command=self._install).pack(side="right")

    # 获取默认安装目录
    def _default_install_dir(self) -> Path:
        local_app_data = os.environ.get("LOCALAPPDATA")
        if local_app_data:
            return Path(local_app_data) / APP_NAME
        return Path.home() / APP_NAME

    # 选择安装目录
    def _choose_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择安装目录", initialdir=self.install_dir_var.get())
        if selected:
            self.install_dir_var.set(selected)

    # 执行安装
    def _install(self) -> None:
        try:
            source_exe = self._resource_path(EXE_NAME)
            if not source_exe.is_file():
                messagebox.showerror("安装失败", "安装包中没有找到主程序")
                return

            install_dir = Path(self.install_dir_var.get()).expanduser()
            install_dir.mkdir(parents=True, exist_ok=True)
            target_exe = install_dir / EXE_NAME
            shutil.copy2(source_exe, target_exe)

            if self.shortcut_var.get():
                self._create_desktop_shortcut(target_exe)

            messagebox.showinfo("安装完成", f"已安装到：{install_dir}")
            self.root.destroy()
        except Exception as exc:
            messagebox.showerror("安装失败", str(exc))

    # 获取打包后的资源路径
    def _resource_path(self, name: str) -> Path:
        base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
        return base_dir / name

    # 创建桌面快捷方式
    def _create_desktop_shortcut(self, target_exe: Path) -> None:
        desktop = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"
        shortcut_path = desktop / "Word 页眉页脚清理工具.lnk"
        shortcut_path_text = self._powershell_quote(str(shortcut_path))
        target_text = self._powershell_quote(str(target_exe))
        working_dir_text = self._powershell_quote(str(target_exe.parent))
        script = (
            "$shell = New-Object -ComObject WScript.Shell; "
            f"$shortcut = $shell.CreateShortcut({shortcut_path_text}); "
            f"$shortcut.TargetPath = {target_text}; "
            f"$shortcut.WorkingDirectory = {working_dir_text}; "
            "$shortcut.Save()"
        )
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            check=True,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )

    # 转义 PowerShell 单引号字符串
    def _powershell_quote(self, text: str) -> str:
        return "'" + text.replace("'", "''") + "'"


# 程序入口
def main() -> int:
    if "--self-test" in sys.argv:
        source_exe = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / EXE_NAME
        if source_exe.is_file():
            print("SETUP SELF TEST PASS")
            return 0
        print("SETUP SELF TEST FAIL")
        return 1

    root = Tk()
    InstallerApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
