import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image
from pptx import Presentation
from pptx.util import Inches

from presenter_gui import PresenterApp


def image_bytes(size, fmt='PNG'):
    out = io.BytesIO()
    Image.new('RGB', size, 'red').save(out, fmt)
    return out.getvalue()


class Var:
    def __init__(self, value): self.value = value
    def get(self): return self.value


class Response:
    def __init__(self, data=None):
        self.status_code = 200 if data else 404
        self.content = data or b''


class ImageQualityTests(unittest.TestCase):
    def setUp(self):
        self.app = PresenterApp.__new__(PresenterApp)
        self.app.local_img_folder = Var('')
        self.app.log = lambda message: None

    def test_uses_higher_resolution_url_instead_of_excel_preview_and_preserves_bytes_in_pptx(self):
        preview = image_bytes((100, 80))
        original = image_bytes((1000, 1000), 'JPEG')
        with patch('presenter_gui.requests.get', side_effect=lambda url, **kwargs: Response(original if url.endswith('.jpg') else None)):
            data, source = self.app.resolve_product_image_bytes({'sku': '123'}, {(1, 3): preview}, 2)
        self.assertEqual(data, original)
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        self.app.add_image_bytes_to_quadrant(slide, data, 0, 0, Inches(5), Inches(3))
        out = io.BytesIO()
        prs.save(out)
        import zipfile
        with zipfile.ZipFile(io.BytesIO(out.getvalue())) as archive:
            media = [n for n in archive.namelist() if n.startswith('ppt/media/')]
            self.assertEqual(archive.read(media[0]), original)

    def test_preserves_better_local_image_when_url_is_lower_resolution(self):
        original = image_bytes((1400, 1200))
        with tempfile.TemporaryDirectory() as directory:
            Path(directory, '123.png').write_bytes(original)
            self.app.local_img_folder = Var(directory)
            with patch('presenter_gui.requests.get', side_effect=lambda url, **kwargs: Response(image_bytes((400, 400)) if url.endswith('.jpg') else None)):
                data, source = self.app.resolve_product_image_bytes({'sku': '123'}, {(1, 3): image_bytes((100, 100))}, 2)
        self.assertEqual(data, original)

    def test_webp_image_is_converted_losslessly_for_powerpoint(self):
        original = image_bytes((250, 150), 'WEBP')
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        self.app.add_image_bytes_to_quadrant(slide, original, 0, 0, Inches(5), Inches(3))
        self.assertEqual(len(slide.shapes), 1)
        inserted = Image.open(io.BytesIO(slide.shapes[0].image.blob))
        self.assertEqual(inserted.size, (250, 150))
        self.assertEqual(inserted.format, 'PNG')

    def test_does_not_use_another_products_adjacent_embedded_image(self):
        with patch('presenter_gui.requests.get', return_value=Response()):
            data, source = self.app.resolve_product_image_bytes({'sku': '123'}, {(2, 3): image_bytes((100, 100))}, 2)
        self.assertIsNone(data)


if __name__ == '__main__':
    unittest.main()
