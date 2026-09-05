import unittest
from io import BytesIO
from unittest.mock import patch

import openpyxl
from openpyxl.styles import PatternFill

import tdm.classinfo as classinfo
from tdm.sparse_worksheet import named_rows


class TemporaryClassInfoTests(unittest.TestCase):
    def make_workbook(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "반 정보"
        ws.append(["반명", "선생님명", "요일", "시간", "모의고사 응시여부"])
        ws.append(["Remove", "Teacher1", "월", 18, "Y"])
        ws["A55"] = "Keep"
        ws["B55"] = "Teacher2"
        ws["E55"] = "Y"
        ws.row_dimensions[55].hidden = True
        ws["J1048576"].fill = PatternFill(fill_type="solid", fgColor="FFFF00")
        self.addCleanup(wb.close)
        return wb

    def create_temp(self, wb, selected):
        output = BytesIO()
        with patch.object(classinfo, "make_backup_file") as backup, \
             patch.object(classinfo, "open", return_value=wb), \
             patch.object(classinfo, "save_to_temp", side_effect=lambda book: book.save(output)), \
             patch.object(wb, "close", wraps=wb.close) as close:
            classinfo.make_temp_file_for_update(selected)
            backup.assert_called_once()
            close.assert_called_once()
        output.seek(0)
        restored = openpyxl.load_workbook(output)
        self.addCleanup(restored.close)
        self.assertLess(len(restored.active._cells), 1000)
        return restored.active

    def test_append_after_last_remaining_class_with_inflated_dimension(self):
        ws = self.create_temp(self.make_workbook(), ["Keep", "Zulu", "Alpha"])
        self.assertEqual(named_rows(ws, 1), [(54, "Keep"), (55, "Alpha"), (56, "Zulu")])
        self.assertEqual(ws["B54"].value, "Teacher2")
        self.assertEqual(ws["E54"].value, "Y")
        self.assertEqual(ws["A55"].border, classinfo.BORDER_ALL)
        self.assertEqual(ws["J1048575"].fill.fgColor.rgb, "00FFFF00")

    def test_no_new_classes(self):
        ws = self.create_temp(self.make_workbook(), ["Keep"])
        self.assertEqual(named_rows(ws, 1), [(54, "Keep")])

    def test_replace_all_classes_starts_at_row_two(self):
        ws = self.create_temp(self.make_workbook(), ["New"])
        self.assertEqual(named_rows(ws, 1), [(2, "New")])

    def test_remove_all_classes(self):
        ws = self.create_temp(self.make_workbook(), [])
        self.assertEqual(named_rows(ws, 1), [])
        self.assertEqual(ws["A1"].value, "반명")

    def test_more_than_one_hundred_classes_are_not_truncated(self):
        selected = [f"Class{i:03}" for i in range(120)]
        ws = self.create_temp(self.make_workbook(), selected)
        self.assertEqual([name for _, name in named_rows(ws, 1)], selected)


if __name__ == "__main__":
    unittest.main()
