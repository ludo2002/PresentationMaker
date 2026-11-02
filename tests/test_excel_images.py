import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image

from presenter_gui import PresenterApp


def png(color):
    out = io.BytesIO()
    Image.new('RGB', (20, 20), color).save(out, format='PNG')
    return out.getvalue()


class Var:
    def get(self): return ''


class ExcelImageTests(unittest.TestCase):
    def setUp(self):
        self.app = PresenterApp.__new__(PresenterApp)
        self.app.log = lambda message: None
        self.app.local_img_folder = Var()

    def test_active_sheet_image_is_not_overwritten_by_other_sheet(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / 'book.xlsx'
            book = Workbook()
            active = book.active
            active.title = 'active'
            active.add_image(XLImage(io.BytesIO(png('red'))), 'D2')
            other = book.create_sheet('other')
            other.add_image(XLImage(io.BytesIO(png('blue'))), 'D2')
            book.save(path)
            result = self.app.extract_embedded_images(str(path))
            image = Image.open(io.BytesIO(result[(1, 3)]))
            self.assertEqual(image.getpixel((0, 0)), (255, 0, 0))

    def test_selected_sheet_image_is_extracted(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / 'book.xlsx'
            book = Workbook()
            active = book.active
            active.title = 'active'
            active.add_image(XLImage(io.BytesIO(png('red'))), 'D2')
            other = book.create_sheet('other')
            other.add_image(XLImage(io.BytesIO(png('blue'))), 'D2')
            book.save(path)

            result = self.app.extract_embedded_images(str(path), sheet_name='other')

            image = Image.open(io.BytesIO(result[(1, 3)]))
            self.assertEqual(image.getpixel((0, 0)), (0, 0, 255))

    def test_real_workbook_image_hyperlink_is_extracted_for_active_row(self):
        path = Path('/home/motolisto/ocbakup/workspace/PresentationMaker/UNDERQ227CLZV_PARA _PRE.xlsx')
        if not path.exists(): self.skipTest('Real workbook is not available')
        self.app.extract_embedded_images(str(path))
        self.assertIn('6007140016', self.app.image_links[(1, 3)][0])

    def test_linked_url_is_used_when_different_from_generated_url(self):
        data = png('green')
        self.app.image_links = {(1, 3): ['https://example.com/original.png']}
        class Response:
            status_code = 200
            content = data
        with patch('presenter_gui.requests.get', side_effect=lambda url, **kw: Response() if url == 'https://example.com/original.png' else type('R', (), {'status_code': 404})()):
            selected, source = self.app.resolve_product_image_bytes({'sku': '123'}, {}, 2)
        self.assertEqual(selected, data)

if __name__ == '__main__': unittest.main()
