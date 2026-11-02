"""需要真实分页或旧版 .doc 时使用桌面 Word/WPS，始终在副本上操作。"""

from contextlib import contextmanager
import ctypes
import os
from pathlib import Path
import re
import shutil
import tempfile


INLINE_PICTURE_TYPES = {3, 4, 7, 8}
FLOATING_PICTURE_TYPES = {11, 13, 28, 29}


@contextmanager
def office_application():
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError as exc:
        raise RuntimeError('此功能需要 Windows 和 pywin32；源码运行请先安装 requirements.txt') from exc

    # pywin32 会合并主线程上的重复初始化；直接调用系统接口以成对计数，
    # 避免清理本次任务时断开调用方已有的 COM 对象。
    com = ctypes.OleDLL('ole32')
    com.CoUninitialize.restype = None
    com.CoInitializeEx(None, 2)  # COINIT_APARTMENTTHREADED
    application = None
    try:
        for program in ('Word.Application', 'KWPS.Application'):
            try:
                application = DispatchEx(program)
                application.Visible = False
                application.DisplayAlerts = 0
                application.AutomationSecurity = 3  # 不运行文档宏
                application.Options.UpdateLinksAtOpen = False
                break
            except Exception:
                if application is not None:
                    try:
                        application.Quit(0)
                    except Exception:
                        pass
                application = None
        if application is None:
            raise RuntimeError('此功能需要安装 Microsoft Word 或支持 COM 自动化的 WPS 文字')
        yield application
    finally:
        if application is not None:
            try:
                application.Quit(0)
            except Exception:
                pass
        application = None
        com.CoUninitialize()


def page_number(text_range):
    position = text_range.Duplicate
    position.Collapse(1)  # wdCollapseStart，避免图片后的分页符影响判断
    return position.Information(3)  # 物理页数，不受分节重新编号影响


def pictures_in_shape(shape):
    if shape.Type in FLOATING_PICTURE_TYPES:
        return [shape]
    if shape.Type in (6, 20):  # 组合和绘图画布：保留其中的文字、线条等
        children = shape.GroupItems if shape.Type == 6 else shape.CanvasItems
        return [picture for index in range(1, children.Count + 1)
                for picture in pictures_in_shape(children.Item(index))]
    if shape.Type == 17:  # 文本框中的嵌入图片
        return [picture for picture in shape.TextFrame.TextRange.InlineShapes
                if picture.Type in INLINE_PICTURE_TYPES]
    return []


def remove_last_page_images(document):
    document.Repaginate()
    last_page = document.ComputeStatistics(2)  # wdStatisticPages
    pictures = []
    for picture in document.InlineShapes:
        if picture.Type in INLINE_PICTURE_TYPES and page_number(picture.Range) == last_page:
            pictures.append(picture)
    for shape in document.Shapes:
        if page_number(shape.Anchor) == last_page:
            pictures.extend(pictures_in_shape(shape))

    # 固定原始末页的目标集合，删除过程中的重新分页不能扩大删除范围。
    for picture in reversed(pictures):
        picture.Delete()
    return len(pictures)


def write_header_footer(document, kind, text, alignment):
    alignment_value = {'left': 0, 'center': 1, 'right': 2}[alignment]
    for section in document.Sections:
        collection = section.Headers if kind == 'header' else section.Footers
        for index in (1, 2, 3):  # 默认页、首页、偶数页
            part = collection.Item(index)
            part.LinkToPrevious = False
            for shape_index in range(part.Shapes.Count, 0, -1):
                part.Shapes.Item(shape_index).Delete()
            part.Range.Text = ''
            cursor = part.Range.Duplicate
            cursor.Collapse(1)
            for segment in re.split(r'(\{PAGE\}|\{NUMPAGES\})', text):
                if segment in ('{PAGE}', '{NUMPAGES}'):
                    field = cursor.Fields.Add(cursor, -1, segment[1:-1], False)
                    cursor.SetRange(field.Result.End + 1, field.Result.End + 1)
                elif segment:
                    cursor.InsertAfter(segment)
                    cursor.Collapse(0)
            part.Range.ParagraphFormat.Alignment = alignment_value
            part.Range.ParagraphFormat.Borders.Enable = False


def process_with_office(source: Path, output: Path, mode: str,
                        header_text=None, footer_text=None,
                        header_alignment='left', footer_alignment='right'):
    """返回 (是否改变, 说明)。保存、关闭成功后才用副本替换目标文件。"""
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.word-cleaner-', dir=output.parent) as folder:
        working_copy = Path(folder) / source.name
        shutil.copyfile(source, working_copy)
        with office_application() as application:
            document = None
            try:
                document = application.Documents.Open(
                    str(working_copy), ConfirmConversions=False, ReadOnly=False,
                    AddToRecentFiles=False, PasswordDocument='', PasswordTemplate='',
                    WritePasswordDocument='', WritePasswordTemplate='',
                    Visible=False, OpenAndRepair=False, NoEncodingDialog=True)
                if document.ReadOnly or document.ProtectionType != -1:
                    raise RuntimeError('文档为只读或受保护，无法修改')
                tracked = document.TrackRevisions
                document.TrackRevisions = False
                if mode == 'last_page_images':
                    count = remove_last_page_images(document)
                    if not count:
                        return False, '最后一页正文中没有图片'
                    message = f'已删除最后一页正文中的 {count} 张图片'
                else:
                    if header_text is not None:
                        write_header_footer(document, 'header', header_text, header_alignment)
                    if footer_text is not None:
                        write_header_footer(document, 'footer', footer_text, footer_alignment)
                    message = '页眉页脚处理完成（保留 .doc 格式）'
                document.TrackRevisions = tracked
                document.Save()
                document.Close(0)
                document = None
            finally:
                if document is not None:
                    document.Close(0)
        os.replace(working_copy, output)
    return True, message
