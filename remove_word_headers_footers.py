#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import os
import posixpath
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Callable, Iterable
from xml.etree import ElementTree as ET


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
WORD_TYPE_ATTR = f"{{{WORD_NS}}}type"
HEADER_REL_TYPE = f"{OFFICE_REL_NS}/header"
FOOTER_REL_TYPE = f"{OFFICE_REL_NS}/footer"
HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
WORD_EXTENSIONS = {".docx", ".docm", ".dotx", ".dotm"}
# 支持的页眉页脚对齐方式
ALIGNMENT_VALUES = {"left", "center", "right"}
HEADER_FOOTER_CONFIGS = {
    "header": {
        "xml_tag": "hdr",
        "reference_tag": "headerReference",
        "part_prefix": "header",
        "rel_id_prefix": "rIdHeader",
        "rel_type": HEADER_REL_TYPE,
        "content_type": HEADER_CONTENT_TYPE,
    },
    "footer": {
        "xml_tag": "ftr",
        "reference_tag": "footerReference",
        "part_prefix": "footer",
        "rel_id_prefix": "rIdFooter",
        "rel_type": FOOTER_REL_TYPE,
        "content_type": FOOTER_CONTENT_TYPE,
    },
}


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


# 生成只包含指定文本的页眉或页脚 XML
def build_header_footer_xml(kind: str, text_value: str, alignment: str) -> bytes:
    if alignment not in ALIGNMENT_VALUES:
        raise ValueError(f"不支持的对齐方式：{alignment}")
    wrapper = ET.Element(f"{{{WORD_NS}}}{HEADER_FOOTER_CONFIGS[kind]['xml_tag']}")
    paragraph = ET.SubElement(wrapper, f"{{{WORD_NS}}}p")
    paragraph_properties = ET.SubElement(paragraph, f"{{{WORD_NS}}}pPr")
    ET.SubElement(paragraph_properties, f"{{{WORD_NS}}}jc", {f"{{{WORD_NS}}}val": alignment})
    run = ET.SubElement(paragraph, f"{{{WORD_NS}}}r")
    text = ET.SubElement(run, f"{{{WORD_NS}}}t")
    text.text = text_value
    if text_value != text_value.strip() or "\n" in text_value:
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return xml_to_bytes(wrapper)


# 获取部件自己的关系文件路径
def build_part_relationship_path(part_name: str) -> str:
    clean_part = part_name.lstrip("/")
    rel_path = PurePosixPath(clean_part)
    return str(rel_path.parent / "_rels" / f"{rel_path.name}.rels")


# 把部件关系中的相对目标转换成文档包内路径
def resolve_relationship_target(relationship_path: str, target: str) -> str:
    clean_target = target.replace("\\", "/")
    if clean_target.startswith("/"):
        return posixpath.normpath(clean_target.lstrip("/"))

    rel_path = PurePosixPath(relationship_path)
    owner_directory = rel_path.parent.parent
    return posixpath.normpath(str(owner_directory / clean_target))


# 收集替换页眉页脚后不再使用的关系文件和图片
def collect_obsolete_part_files(source_zip: zipfile.ZipFile, updated_parts: set[str]) -> set[str]:
    names = set(source_zip.namelist())
    removed_relationships = {
        relationship_path
        for part in updated_parts
        if (relationship_path := build_part_relationship_path(part)) in names
    }
    candidate_images: set[str] = set()

    for relationship_path in removed_relationships:
        try:
            root = parse_xml(source_zip.read(relationship_path))
        except (KeyError, ET.ParseError):
            continue
        for relationship in root:
            if relationship.attrib.get("TargetMode") == "External":
                continue
            if not relationship.attrib.get("Type", "").endswith("/image"):
                continue
            target = relationship.attrib.get("Target", "")
            if target:
                candidate_images.add(resolve_relationship_target(relationship_path, target))

    remaining_targets: set[str] = set()
    for relationship_path in names:
        if not relationship_path.endswith(".rels") or relationship_path in removed_relationships:
            continue
        try:
            root = parse_xml(source_zip.read(relationship_path))
        except (KeyError, ET.ParseError):
            return removed_relationships
        for relationship in root:
            if relationship.attrib.get("TargetMode") == "External":
                continue
            target = relationship.attrib.get("Target", "")
            if target:
                remaining_targets.add(resolve_relationship_target(relationship_path, target))

    return removed_relationships | {
        image_path
        for image_path in candidate_images
        if image_path in names and image_path not in remaining_targets
    }


# 把关系文件里的 Target 转成 docx 内部路径
def relationship_target_to_zip_path(target: str) -> str:
    target_path = PurePosixPath(target)
    if target_path.is_absolute():
        return str(target_path).lstrip("/")
    return str(PurePosixPath("word") / target_path)


# 收集关系文件中已经存在的页眉或页脚文件路径
def collect_header_footer_targets(rels_xml: bytes, kind: str) -> set[str]:
    root = parse_xml(rels_xml)
    targets: set[str] = set()
    rel_type = HEADER_FOOTER_CONFIGS[kind]["rel_type"]

    # 关系的 Target 指向文档包里的实际页眉或页脚文件
    for relationship in root:
        if relationship.attrib.get("Type", "") == rel_type:
            rel_target = relationship.attrib.get("Target", "")
            if rel_target:
                targets.add(relationship_target_to_zip_path(rel_target))

    return targets


# 生成不和现有文件冲突的页眉或页脚文件路径
def build_header_footer_part_name(kind: str, existing_names: set[str]) -> str:
    part_prefix = HEADER_FOOTER_CONFIGS[kind]["part_prefix"]
    index = 1
    while True:
        candidate = f"word/{part_prefix}{index}.xml"
        if candidate not in existing_names:
            return candidate
        index += 1


# 生成不和现有关系冲突的关系编号
def build_relationship_id(kind: str, root: ET.Element) -> str:
    rel_id_prefix = HEADER_FOOTER_CONFIGS[kind]["rel_id_prefix"]
    existing_ids = {child.attrib.get("Id", "") for child in root}
    if rel_id_prefix not in existing_ids:
        return rel_id_prefix

    index = 1
    while f"{rel_id_prefix}{index}" in existing_ids:
        index += 1
    return f"{rel_id_prefix}{index}"


# 给关系文件新增页眉或页脚关系
def add_header_footer_relationship(rels_xml: bytes, kind: str, rel_id: str, part_name: str) -> bytes:
    root = parse_xml(rels_xml)
    ET.SubElement(
        root,
        f"{{{REL_NS}}}Relationship",
        {
            "Id": rel_id,
            "Type": HEADER_FOOTER_CONFIGS[kind]["rel_type"],
            "Target": PurePosixPath(part_name).name,
        },
    )
    return xml_to_bytes(root)


# 确保正文分节设置引用指定页眉或页脚
def ensure_header_footer_reference(document_xml: bytes, kind: str, rel_id: str) -> tuple[bytes, bool]:
    root = parse_xml(document_xml)
    body = root.find(f".//{{{WORD_NS}}}body")
    if body is None:
        return document_xml, False

    changed = False
    sections = list(root.iter(f"{{{WORD_NS}}}sectPr"))
    reference_tag = HEADER_FOOTER_CONFIGS[kind]["reference_tag"]
    if not sections:
        sections = [ET.SubElement(body, f"{{{WORD_NS}}}sectPr")]
        changed = True

    # 每个分节都要有默认引用，避免只有部分页面生效
    for section in sections:
        default_reference = None
        for child in section:
            if child.tag.rsplit("}", 1)[-1] == reference_tag and child.attrib.get(WORD_TYPE_ATTR, "default") == "default":
                default_reference = child
                break
        if default_reference is None:
            default_reference = ET.Element(f"{{{WORD_NS}}}{reference_tag}", {WORD_TYPE_ATTR: "default", REL_ATTR: rel_id})
            section.insert(0, default_reference)
            changed = True
        elif default_reference.attrib.get(REL_ATTR) != rel_id:
            default_reference.set(REL_ATTR, rel_id)
            changed = True

    return xml_to_bytes(root), changed


# 确保内容类型表声明了页眉或页脚文件
def ensure_content_type_override(content_types_xml: bytes, kind: str, part_name: str) -> tuple[bytes, bool]:
    root = parse_xml(content_types_xml)
    normalized_part_name = f"/{part_name.lstrip('/')}"

    for child in root:
        if child.tag.rsplit("}", 1)[-1] == "Override" and child.attrib.get("PartName") == normalized_part_name:
            return content_types_xml, False

    ET.SubElement(
        root,
        f"{{{CONTENT_TYPES_NS}}}Override",
        {
            "PartName": normalized_part_name,
            "ContentType": HEADER_FOOTER_CONFIGS[kind]["content_type"],
        },
    )
    return xml_to_bytes(root), True


# 准备页眉或页脚修改所需的包内文件替换
def prepare_header_footer_update(
    document_xml: bytes,
    rels_xml: bytes,
    names: set[str],
    replacements: dict[str, bytes],
    new_files: dict[str, bytes],
    kind: str,
    text_value: str,
    alignment: str,
    read_file: Callable[[str], bytes],
) -> tuple[bytes, bytes, set[str], bool]:
    part_bytes = build_header_footer_xml(kind, text_value, alignment)
    part_targets = collect_header_footer_targets(rels_xml, kind)
    changed = False

    # 已存在时替换原有文件，不存在时补齐文件、关系和正文引用
    if part_targets:
        for part_target in part_targets:
            if part_target not in names or read_file(part_target) != part_bytes:
                changed = True
            if part_target in names:
                replacements[part_target] = part_bytes
            else:
                new_files[part_target] = part_bytes
    else:
        known_names = names | set(new_files)
        part_name = build_header_footer_part_name(kind, known_names)
        rels_root = parse_xml(rels_xml)
        rel_id = build_relationship_id(kind, rels_root)
        document_xml, document_changed = ensure_header_footer_reference(document_xml, kind, rel_id)
        rels_xml = add_header_footer_relationship(rels_xml, kind, rel_id, part_name)
        new_files[part_name] = part_bytes
        part_targets = {part_name}
        changed = True or document_changed

    return document_xml, rels_xml, part_targets, changed


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
        removed_paths.add(build_part_relationship_path(clean_part))
    return removed_paths


# 为输出文件生成不重复的路径
def build_output_path(source: Path, output_dir: Path | None, mode: str) -> Path:
    suffix_map = {
        "header": "no_header",
        "footer": "no_footer",
        "both": "no_header_footer",
        "set_header": "header_updated",
        "set_footer": "footer_updated",
        "set_both": "header_footer_updated",
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
    header_text: str | None = None,
    footer_text: str | None = None,
    header_alignment: str = "left",
    footer_alignment: str = "right",
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
            new_files: dict[str, bytes] = {}
            removed_parts: set[str] = set()
            obsolete_paths: set[str] = set()
            text_changed = False
            document_bytes = zin.read("word/document.xml")
            rels_bytes = zin.read("word/_rels/document.xml.rels")

            if header_text is None and footer_text is None:
                document_bytes, _removed_ids, document_changed = remove_section_references(
                    document_bytes,
                    mode,
                )
                rels_bytes, rel_removed_parts, rels_changed = remove_document_relationships(
                    rels_bytes,
                    mode,
                )
                replacements["word/document.xml"] = document_bytes
                replacements["word/_rels/document.xml.rels"] = rels_bytes
                removed_parts.update(rel_removed_parts)
            else:
                document_changed = False
                rels_changed = False
                all_part_targets: set[str] = set()
                updated_part_targets: set[str] = set()

                if header_text is not None:
                    old_document_bytes = document_bytes
                    old_rels_bytes = rels_bytes
                    document_bytes, rels_bytes, header_targets, header_changed = prepare_header_footer_update(
                        document_bytes,
                        rels_bytes,
                        names,
                        replacements,
                        new_files,
                        "header",
                        header_text,
                        header_alignment,
                        zin.read,
                    )
                    document_changed = document_changed or document_bytes != old_document_bytes
                    rels_changed = rels_changed or rels_bytes != old_rels_bytes
                    text_changed = text_changed or header_changed
                    all_part_targets.update(header_targets)
                    updated_part_targets.update(header_targets)

                if footer_text is not None:
                    old_document_bytes = document_bytes
                    old_rels_bytes = rels_bytes
                    document_bytes, rels_bytes, footer_targets, footer_changed = prepare_header_footer_update(
                        document_bytes,
                        rels_bytes,
                        names,
                        replacements,
                        new_files,
                        "footer",
                        footer_text,
                        footer_alignment,
                        zin.read,
                    )
                    document_changed = document_changed or document_bytes != old_document_bytes
                    rels_changed = rels_changed or rels_bytes != old_rels_bytes
                    text_changed = text_changed or footer_changed
                    all_part_targets.update(footer_targets)
                    updated_part_targets.update(footer_targets)

                obsolete_paths.update(collect_obsolete_part_files(zin, updated_part_targets))
                text_changed = text_changed or bool(obsolete_paths)

                replacements["word/document.xml"] = document_bytes
                replacements["word/_rels/document.xml.rels"] = rels_bytes

            if "[Content_Types].xml" in names:
                if header_text is None and footer_text is None:
                    content_types_bytes, content_types_changed = remove_content_type_overrides(
                        zin.read("[Content_Types].xml"),
                        removed_parts,
                    )
                else:
                    content_types_bytes = zin.read("[Content_Types].xml")
                    content_types_changed = False
                    for part_target in all_part_targets:
                        kind = "header" if PurePosixPath(part_target).name.startswith("header") else "footer"
                        content_types_bytes, current_changed = ensure_content_type_override(
                            content_types_bytes,
                            kind,
                            part_target,
                        )
                        content_types_changed = content_types_changed or current_changed
                replacements["[Content_Types].xml"] = content_types_bytes
            else:
                content_types_changed = False

            changed = (
                document_changed
                or rels_changed
                or content_types_changed
                or text_changed
                or bool(removed_parts)
                or bool(obsolete_paths)
            )
            if not changed:
                if header_text is None and footer_text is None:
                    return ProcessResult(source, None, False, "没有找到需要移除的页眉页脚")
                return ProcessResult(source, None, False, "页眉页脚内容已经一致")

            removed_paths = collect_removed_paths(removed_parts) | obsolete_paths

            # 复制原文档包，跳过要删除的页眉页脚文件，并替换被修改的 XML
            with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in removed_paths:
                        continue
                    if item.filename in replacements:
                        zout.writestr(item, replacements[item.filename])
                    else:
                        zout.writestr(item, zin.read(item.filename))
                for filename, data in new_files.items():
                    if filename not in names:
                        zout.writestr(filename, data)

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
        choices=("header", "footer", "both", "set_header", "set_footer", "set_both"),
        default="both",
        help="header 只去页眉，footer 只去页脚，both 两者都去，set_header 修改或新增页眉，set_footer 修改或新增页脚，set_both 两者都修改，默认 both",
    )
    parser.add_argument(
        "--header-text",
        help="要写入的页眉内容，填写后会修改或新增页眉",
    )
    parser.add_argument(
        "--footer-text",
        help="要写入的页脚内容，填写后会修改或新增页脚",
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

    mode = args.mode
    if args.header_text is not None and args.footer_text is not None:
        mode = "set_both"
    elif args.header_text is not None:
        mode = "set_header"
    elif args.footer_text is not None:
        mode = "set_footer"

    if mode == "set_header" and args.header_text is None:
        parser.error("使用 set_header 模式时必须填写 --header-text")
    if mode == "set_footer" and args.footer_text is None:
        parser.error("使用 set_footer 模式时必须填写 --footer-text")
    if mode == "set_both" and (args.header_text is None or args.footer_text is None):
        parser.error("使用 set_both 模式时必须同时填写 --header-text 和 --footer-text")

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
            mode,
            output_dir,
            in_place,
            header_text=args.header_text,
            footer_text=args.footer_text,
        )
        for source in targets
    ]
    return print_results(results)


if __name__ == "__main__":
    raise SystemExit(main())
