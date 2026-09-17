import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock
from functools import lru_cache
from PIL import Image
import numpy as np

import image_matcher
import dpi_support
from navigation_regression import functions


class DpiTests(unittest.TestCase):
    def test_actual_image_matching_at_common_scales(self):
        rng = np.random.default_rng(17)
        base = Image.fromarray(rng.integers(0, 255, (24, 64, 3), dtype=np.uint8))
        with tempfile.TemporaryDirectory() as folder:
            path = str(Path(folder) / 'menu.png')
            base.save(path)
            for scale in (1, 1.25, 1.5, 1.75, 2, 2.5, 3):
                with self.subTest(scale=scale):
                    template = base.resize((round(64 * scale), round(24 * scale)), Image.Resampling.LANCZOS)
                    screen = Image.new('RGB', (300, 150), '#20252B')
                    screen.paste(template, (31, 27))
                    box, found_scale = image_matcher.find(path, screen, confidence=0.98, preferred=scale)
                    self.assertIsNotNone(box)
                    self.assertEqual((box.left, box.top), (31, 27))
                    self.assertAlmostEqual(found_scale, scale)

    def test_capture_dpi_ratios_include_downscaling(self):
        scales = dpi_support.candidate_scales(1)
        self.assertTrue(any(abs(s - 0.5) < 0.01 for s in scales))
        self.assertTrue(any(abs(s - 1 / 1.75) < 0.01 for s in scales))

    def test_ocr_coordinates_restored_after_normalization(self):
        box = Mock(left=10, top=20, width=30, height=10, confidence=90, text='Monitoring')
        clock = Mock()
        clock.monotonic.return_value = 0
        mod = functions('ocr_helpers.py', 'find_text_in_image',
                        _tesseract_available=True, _scan_for_matches=Mock(return_value=[box]),
                        Image=Image, control=Mock(), time=clock)
        result = mod.find_text_in_image(Image.new('RGB', (600, 400)), 'Monitoring', source_scale=2)
        self.assertEqual((result.left, result.top, result.width, result.height), (20, 40, 60, 20))
        processed = mod._scan_for_matches.call_args.args[0]
        self.assertEqual(processed.size, (300, 200))
