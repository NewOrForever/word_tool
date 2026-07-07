#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import unittest

from remove_word_headers_footers import build_parser
from word_header_footer_gui import MODE_LABELS


class RemovedImageWatermarkModeTests(unittest.TestCase):
    # 验证界面不再显示图片水印处理选项
    def test_gui_mode_labels_do_not_include_image_watermark(self) -> None:
        self.assertNotIn("image_watermark", MODE_LABELS)

    # 验证命令行不再支持图片水印处理模式
    def test_parser_choices_do_not_include_image_watermark(self) -> None:
        parser = build_parser()
        mode_action = next(action for action in parser._actions if action.dest == "mode")

        self.assertNotIn("image_watermark", mode_action.choices)


if __name__ == "__main__":
    unittest.main()
