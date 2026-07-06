#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Iterable
from xml.etree import ElementTree as ET


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
HEADER_REL_TYPE = f"{OFFICE_REL_NS}/header"
FOOTER_REL_TYPE = f"{OFFICE_REL_NS}/footer"
WORD_EXTENSIONS = {".docx", ".docm", ".dotx", ".dotm"}


ET.register_namespace("w", WORD_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("", REL_NS)


# 设置控制台编码，避免 Windows 终端中文输出乱码
def configure_console_encoding() -> None:
    if os.name == "nt":
        os.system("chcp 65001 >nul")

    # 固定 Python 输出编码，确保中文按 UTF-8 输出
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")


@dataclass
class ProcessResult:
    # 原始文件路径
    source: Path
    # 输出文件路径
    output: Path | None
    # 是否处理成功
    changed: bool
    # 当前文件的处理说明
    message: str


# 解析 XML 文件内容
def parse_xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


# 将 XML 写回为带声明的字节内容
def xml_to_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


# 判断当前目标是否需要移除
def should_remove(kind: str, mode: str) -> bool:
    return mode == "both" or mode == kind


# 把关系文件里的 Target 转成 docx 内部路径
def relationship_target_to_zip_path(target: str) -> str:
    target_path = PurePosixPath(target)
    if target_path.is_absolute():
        return str(target_path).lstrip("/")
    return str(PurePosixPath("word") / target_path)


# 移除 document.xml 中的页眉或页脚引用
def remove_section_references(document_xml: bytes, mode: str) -> tuple[bytes, set[str], bool]:
    root = parse_xml(document_xml)
    removed_ids: set[str] = set()
    changed = False

    # 遍历每个分节设置，删除对应的页眉页脚引用
    for section in root.iter(f"{{{WORD_NS}}}sectPr"):
        for child in list(section):
            tag_name = child.tag.rsplit("}", 1)[-1]
            if tag_name == "headerReference" and should_remove("header", mode):
                rel_id = child.attrib.get(REL_ATTR)
                if rel_id:
                    removed_ids.add(rel_id)
                section.remove(child)
                changed = True
            elif tag_name == "footerReference" and should_remove("footer", mode):
                rel_id = child.attrib.get(REL_ATTR)
                if rel_id:
                    removed_ids.add(rel_id)
                section.remove(child)
                changed = True

    return xml_to_bytes(root), removed_ids, changed


# 移除 Word 关系文件中对应的页眉页脚关系
def remove_document_relationships(rels_xml: bytes, mode: str) -> tuple[bytes, set[str], bool]:
    root = parse_xml(rels_xml)
    removed_targets: set[str] = set()
    changed = False

    # 关系文件决定了页眉页脚实际对应的内部文件
    for relationship in list(root):
        rel_type = relationship.attrib.get("Type", "")
        rel_target = relationship.attrib.get("Target", "")
        if rel_type == HEADER_REL_TYPE and should_remove("header", mode):
            removed_targets.add(relationship_target_to_zip_path(rel_target))
            root.remove(relationship)
            changed = True
        elif rel_type == FOOTER_REL_TYPE and should_remove("footer", mode):
            removed_targets.add(relationship_target_to_zip_path(rel_target))
            root.remove(relationship)
            changed = True

    return xml_to_bytes(root), removed_targets, changed


# 移除内容类型表里的页眉页脚声明
def remove_content_type_overrides(content_types_xml: bytes, removed_parts: set[str]) -> tuple[bytes, bool]:
    root = parse_xml(content_types_xml)
    changed = False
    normalized_parts = {f"/{part}" for part in removed_parts}

    # 删除已经移除的页眉页脚文件声明，保持文档包结构干净
    for child in list(root):
        if child.tag.rsplit("}", 1)[-1] != "Override":
            continue
        if child.attrib.get("PartName") in normalized_parts:
            root.remove(child)
            changed = True

    return xml_to_bytes(root), changed


# 收集要跳过写入的新旧文件
def collect_removed_paths(removed_parts: set[str]) -> set[str]:
    removed_paths: set[str] = set()
    for part in removed_parts:
        clean_part = part.lstrip("/")
        removed_paths.add(clean_part)
        rel_path = PurePosixPath(clean_part)
        removed_paths.add(str(rel_path.parent / "_rels" / f"{rel_path.name}.rels"))
    return removed_paths


# 为输出文件生成不重复的路径
def build_output_path(source: Path, output_dir: Path | None, mode: str) -> Path:
    suffix_map = {
        "header": "no_header",
        "footer": "no_footer",
        "both": "no_header_footer",
    }
    target_dir = output_dir if output_dir else source.parent
    candidate = target_dir / f"{source.stem}_{suffix_map[mode]}{source.suffix}"
    index = 1
    while candidate.exists():
        candidate = target_dir / f"{source.stem}_{suffix_map[mode]}_{index}{source.suffix}"
        index += 1
    return candidate


# 处理单个 docx 文件
def process_docx(
    source: Path,
    mode: str,
    output_dir: Path | None,
    in_place: bool,
) -> ProcessResult:
    if source.name.startswith("~$"):
        return ProcessResult(source, None, False, "跳过 Word 临时文件")
    if source.suffix.lower() not in WORD_EXTENSIONS:
        return ProcessResult(source, None, False, "跳过非新版 Word 文件")
    if not zipfile.is_zipfile(source):
        return ProcessResult(source, None, False, "文件不是有效的 docx")

    output = source if in_place else build_output_path(source, output_dir, mode)
    output.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_output = Path(temp_dir) / source.name

        with zipfile.ZipFile(source, "r") as zin:
            names = set(zin.namelist())
            if "word/document.xml" not in names or "word/_rels/document.xml.rels" not in names:
                return ProcessResult(source, None, False, "文件结构不完整")

            replacements: dict[str, bytes] = {}
            removed_parts: set[str] = set()

            document_bytes, _removed_ids, document_changed = remove_section_references(
                zin.read("word/document.xml"),
                mode,
            )
            rels_bytes, rel_removed_parts, rels_changed = remove_document_relationships(
                zin.read("word/_rels/document.xml.rels"),
                mode,
            )
            replacements["word/document.xml"] = document_bytes
            replacements["word/_rels/document.xml.rels"] = rels_bytes
            removed_parts.update(rel_removed_parts)

            if "[Content_Types].xml" in names:
                content_types_bytes, content_types_changed = remove_content_type_overrides(
                    zin.read("[Content_Types].xml"),
                    removed_parts,
                )
                replacements["[Content_Types].xml"] = content_types_bytes
            else:
                content_types_changed = False

            changed = document_changed or rels_changed or content_types_changed or bool(removed_parts)
            if not changed:
                return ProcessResult(source, None, False, "没有找到需要移除的页眉页脚")

            removed_paths = collect_removed_paths(removed_parts)

            # 复制原文档包，跳过要删除的页眉页脚文件，并替换被修改的 XML
            with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in removed_paths:
                        continue
                    if item.filename in replacements:
                        zout.writestr(item, replacements[item.filename])
                    else:
                        zout.writestr(item, zin.read(item.filename))

        shutil.move(str(temp_output), str(output))

    return ProcessResult(source, output, True, "处理完成")


# 根据输入路径收集待处理文件
def collect_targets(target: Path, recursive: bool) -> list[Path]:
    if target.is_file():
        return [target]
    if not target.is_dir():
        raise FileNotFoundError(f"路径不存在：{target}")

    pattern = "**/*" if recursive else "*"
    return sorted(
        path for path in target.glob(pattern)
        if path.is_file()
        and path.suffix.lower() in WORD_EXTENSIONS
    )


# 打印每个文件的处理结果
def print_results(results: Iterable[ProcessResult]) -> int:
    success_count = 0
    skip_count = 0

    for result in results:
        if result.changed and result.output:
            success_count += 1
            print(f"[OK] {result.source} -> {result.output}")
        else:
            skip_count += 1
            print(f"[SKIP] {result.source}：{result.message}")

    print(f"\n完成：成功 {success_count} 个，跳过 {skip_count} 个")
    return 0 if success_count > 0 else 1


# 构建命令行参数
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="批量移除新版 Word 文件的页眉和页脚",
    )
    parser.add_argument("target", help="单个新版 Word 文件路径，或包含新版 Word 文件的目录路径")
    parser.add_argument(
        "--mode",
        choices=("header", "footer", "both"),
        default="both",
        help="header 只去页眉，footer 只去页脚，both 两者都去，默认 both",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="处理目录时递归处理子目录",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="输出目录，不填写则直接覆盖原文件",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="直接覆盖原文件，不填写输出目录时默认就是覆盖",
    )
    return parser


# 程序入口
def main() -> int:
    configure_console_encoding()

    parser = build_parser()
    args = parser.parse_args()

    target = Path(args.target).expanduser().resolve()
    output_dir = args.output_dir.expanduser().resolve() if args.output_dir else None

    if args.in_place and output_dir:
        parser.error("--in-place 和 --output-dir 不能同时使用")

    in_place = args.in_place or output_dir is None

    try:
        targets = collect_targets(target, args.recursive)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    if not targets:
        print("没有找到可处理的 docx 文件", file=sys.stderr)
        return 1

    results = [
        process_docx(
            source,
            args.mode,
            output_dir,
            in_place,
        )
        for source in targets
    ]
    return print_results(results)


if __name__ == "__main__":
    raise SystemExit(main())
