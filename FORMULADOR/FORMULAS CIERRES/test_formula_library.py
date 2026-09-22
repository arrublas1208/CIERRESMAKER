import unittest

from PySide6.QtCore import QEvent, Qt
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import QApplication

import formula_builder


class FormulaLibraryTabTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_equal_key_on_table_starts_fx(self):
        tab = formula_builder.FormulaLibraryTab()
        tab.table.setCurrentCell(0, 1)

        event = QKeyEvent(QEvent.Type.KeyPress, Qt.Key.Key_Equal, Qt.KeyboardModifier.NoModifier, "=")
        handled = tab.eventFilter(tab.table, event)

        self.assertTrue(handled)
        self.assertTrue(tab.fx_active)
        self.assertEqual(tab.fx_edit.text(), "=")

    def test_equal_key_on_viewport_starts_fx(self):
        tab = formula_builder.FormulaLibraryTab()
        tab.table.setCurrentCell(0, 1)

        event = QKeyEvent(QEvent.Type.KeyPress, Qt.Key.Key_Equal, Qt.KeyboardModifier.NoModifier, "=")
        handled = tab.eventFilter(tab.table.viewport(), event)

        self.assertTrue(handled)
        self.assertTrue(tab.fx_active)
        self.assertEqual(tab.fx_edit.text(), "=")


if __name__ == "__main__":
    unittest.main()
