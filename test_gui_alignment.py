#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import unittest

from word_header_footer_gui import alignment_value


class GuiAlignmentTests(unittest.TestCase):
    # 验证界面中的位置文字会转换成文档对齐值
    def test_alignment_labels_map_to_document_values(self) -> None:
        self.assertEqual("left", alignment_value("左边"))
        self.assertEqual("center", alignment_value("居中"))
        self.assertEqual("right", alignment_value("右边"))

    # 验证未知位置不会静默写入文档
    def test_unknown_alignment_label_is_rejected(self) -> None:
        with self.assertRaisesRegex(ValueError, "不支持的位置"):
            alignment_value("顶部")


if __name__ == "__main__":
    unittest.main()
