import unittest

from openpyxl import Workbook

from presenter_gui import PresenterApp


class SheetSelectionTests(unittest.TestCase):
    def test_select_worksheet_returns_requested_sheet(self):
        workbook = Workbook()
        workbook.active.title = "Primera"
        workbook.create_sheet("Otra hoja")

        selected = PresenterApp.select_worksheet(workbook, "Otra hoja")

        self.assertEqual(selected.title, "Otra hoja")

    def test_select_worksheet_defaults_to_active_sheet(self):
        workbook = Workbook()
        workbook.active.title = "Primera"
        workbook.create_sheet("Activa")
        workbook.active = 1

        selected = PresenterApp.select_worksheet(workbook, "")

        self.assertEqual(selected.title, "Activa")


if __name__ == "__main__":
    unittest.main()
