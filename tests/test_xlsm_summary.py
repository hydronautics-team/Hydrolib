import json
from pathlib import Path
import sys
import tempfile
import unittest
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools/xlsm_merge'))
from summary import render

class SummaryTests(unittest.TestCase):
    def test_success_and_early_failure(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertIn('Книгу удалось', render(d, 'success', 'a/b', '1'))
            self.assertIn('не запускалась', render(d, 'skipped', 'a/b', '1'))

    def test_conflict_and_fork_are_clear_and_escaped(self):
        with tempfile.TemporaryDirectory() as d:
            Path(d, 'report.json').write_text(json.dumps({'conflicts': [{'part': 'A', 'field': 'Name', 'ours': '<script>', 'theirs': 'B', 'choice': None}]}))
            text = render(d, 'failure', 'a/b', '3')
            self.assertIn('нужен выбор стороны', text)
            self.assertIn('&lt;script&gt;', text)
            self.assertIn('xlsm-update.yml', text)
            Path(d, 'context.json').write_text(json.dumps({'number': '3', 'fork': True}))
            self.assertIn('PR из форка', render(d, 'failure', 'a/b', '3'))

    def test_unresolved_and_resolved_are_separate(self):
        with tempfile.TemporaryDirectory() as d:
            Path(d, 'report.json').write_text(json.dumps({'conflicts': [
                {'part': 'A', 'base': 'old', 'ours': 'new', 'theirs': 'other', 'choice': None},
                {'part': 'B', 'base': 'old', 'ours': 'new', 'theirs': 'other', 'choice': 'ours'}]}))
            text = render(d, 'failure', 'a/b', '1')
            self.assertIn('Нерешённые конфликты: 1', text)
            self.assertIn('Разрешённые конфликты: 1', text)
            self.assertIn('<th>Исходное</th>', text)

    def test_structural_error_is_visible(self):
        with tempfile.TemporaryDirectory() as d:
            Path(d, 'report.json').write_text(json.dumps({'message': 'Column schema changed', 'conflicts': []}))
            text = render(d, 'failure', 'a/b', '1')
            self.assertIn('Что мешает завершить слияние', text)
            self.assertIn('Column schema changed', text)
