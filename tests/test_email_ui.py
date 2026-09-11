import unittest
from pathlib import Path
from streamlit.testing.v1 import AppTest

class EmailUITests(unittest.TestCase):
    def test_generation_and_stale_downloads(self):
        app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / 'app.py')).run()
        self.assertFalse(app.exception)
        self.assertIn('合庫', app.title[0].value)
        app.session_state['formatter_mode'] = '財金新聞信件排版'
        app.run()
        app.text_area(key='email_source_text').set_value('## 測試標題\n測試內文。').run()
        app.button(key='email_generate').click().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.get('download_button')), 3)
        app.run()
        self.assertEqual(len(app.get('download_button')), 3)
        app.text_input(key='email_subject').set_value('新主旨').run()
        self.assertEqual(len(app.get('download_button')), 0)
        app.button(key='email_generate').click().run()
        self.assertEqual(len(app.get('download_button')), 3)
        app.session_state['formatter_mode'] = 'Word 速報排版'
        app.run()
        self.assertFalse(app.exception)
        self.assertIn('合庫', app.title[0].value)
