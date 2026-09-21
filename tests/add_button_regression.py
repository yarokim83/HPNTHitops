import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from PIL import Image
import pr_add_button as add


class AddButtonTests(unittest.TestCase):
    def test_supplied_toolbar_at_multiple_scales(self):
        with Image.open(Path(__file__).parent / 'pr_toolbar_fixture.png') as image:
            for scale in (1, 1.25, 1.5, 1.75, 2, 2.5, 3):
                with self.subTest(scale=scale):
                    sample = image.resize((round(image.width*scale), round(image.height*scale)), Image.Resampling.LANCZOS)
                    match = add.find_plus(sample, scale)
                    self.assertIsNotNone(match)
                    self.assertAlmostEqual(match[0]/scale, 81.5, delta=1)
                    self.assertAlmostEqual(match[1]/scale, 15.5, delta=1)

    def test_red_minus_and_other_toolbar_icons_are_rejected(self):
        with Image.open(Path(__file__).parent / 'pr_toolbar_fixture.png') as image:
            sample = image.copy()
        sample.paste('#dddddd', (69, 0, 95, 31))
        self.assertIsNone(add.find_plus(sample))

    def test_greyscale_plus_is_not_accepted(self):
        with Image.open(Path(__file__).parent / 'pr_toolbar_fixture.png') as image:
            self.assertIsNone(add.find_plus(image.convert('L').convert('RGB')))

    def test_search_is_restricted_to_client_toolbar_and_restores_negative_origin(self):
        api = Mock()
        api.IsWindowEnabled.return_value = True
        api.IsIconic.return_value = False
        api.ClientToScreen.return_value = (-1800, 50)
        api.GetClientRect.return_value = (0, 0, 1200, 900)
        with patch.object(add, 'win32gui', api), patch.object(add, 'monitor_scale', return_value=1.5), \
             patch.object(add.ImageGrab, 'grab') as grab, patch.object(add, 'find_plus', return_value=(120, 24, .98, 1.5)):
            self.assertEqual(add.locate(1)[0], (-1680, 74))
            grab.assert_called_once_with(bbox=(-1800, 50, -1350, 98), all_screens=True)
