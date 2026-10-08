"""Direct-selection XML: displayed values, stable reads and clipboard isolation."""
from pathlib import Path
import sys
import unittest
from types import SimpleNamespace as NS
from unittest import mock
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path[:0] = [str(ROOT), str(ROOT / 'windows')]
from tabledown_windows import excel_table as table, excel_formula as formula
from tablemark.excel_table import ExcelTableSelection
import test_windows_port as port


def selected(raw=((1, None), ('', ' x ')), display=('1.00', '', '', ' x '), merges=None):
    rows, columns = len(raw), len(raw[0])
    cells = [NS(Text=value, MergeArea=NS(Address=(merges or {}).get(i, f'${chr(65+i%columns)}${1+i//columns}'))) for i, value in enumerate(display)]
    return NS(Worksheet=NS(Name='Sheet1', Parent=NS(Name='Book.xlsx')),
              Address='$A$1:$B$2', CountLarge=rows*columns,
              Areas=NS(Count=1), Rows=NS(Count=rows, CountLarge=rows),
              Columns=NS(Count=columns, CountLarge=columns), Row=1, Column=1,
              Value2=raw, Text=None, MergeCells=None if merges else False, Cells=cells)


class TableReaderTests(unittest.TestCase):
    def read(self, selection, *, cancelled=None, identity=None):
        com = mock.Mock()
        excel = NS(Hwnd=5, Selection=selection)
        client = NS(GetActiveObject=mock.Mock(return_value=excel))
        with mock.patch.object(formula, '_load_com_modules', return_value=(com, client)), mock.patch.object(formula, '_visible_excel_process_ids', return_value={7}), mock.patch.object(formula, '_window_process_id', return_value=7):
            if identity:
                with mock.patch.object(formula, '_selection_identity', side_effect=identity):
                    result = table.read_selected_excel_table(is_cancelled=cancelled)
            else:
                result = table.read_selected_excel_table(is_cancelled=cancelled)
        com.CoInitialize.assert_called_once()
        com.CoUninitialize.assert_called_once()
        return result

    def test_formatted_text_spaces_empty_and_blank(self):
        result = self.read(selected())
        self.assertTrue(result.ok, result.code)
        self.assertEqual(result.selection.values, ('1.00', None, '', ' x '))

    def test_mixed_merge_deduplicated(self):
        result = self.read(selected(merges={0:'$A$1:$B$1', 1:'$A$1:$B$1'}))
        self.assertTrue(result.ok, result.code)
        self.assertEqual(result.selection.merge_areas, ('$A$1:$B$1',))

    def test_partial_merge_rejected(self):
        self.assertEqual(self.read(selected(merges={0:'$A$1:$C$1'})).code, 'partial_merge')

    def test_uniform_text_uses_range_fast_path(self):
        selection = selected(raw=((1,1),(1,1)))
        selection.Text, selection.Cells = '1.00', None
        result = self.read(selection)
        self.assertTrue(result.ok, result.code)
        self.assertEqual(result.selection.values, ('1.00',)*4)

    def test_hash_display_rejected_but_literal_preserved(self):
        self.assertEqual(self.read(selected(display=('###','','',' x '))).code, 'display_overflow')
        result = self.read(selected(raw=(('###',None),('', ' x ')), display=('###','','',' x ')))
        self.assertTrue(result.ok, result.code)
        self.assertEqual(result.selection.values[0], '###')

    def test_limits_and_nonrectangular_selection(self):
        selection = selected()
        selection.CountLarge = 10001
        selection.Rows.CountLarge = selection.Rows.Count = 10001
        selection.Columns.CountLarge = selection.Columns.Count = 1
        self.assertEqual(self.read(selection).code, formula.TOO_LARGE)
        selection.Areas.Count = 2
        self.assertEqual(self.read(selection).code, formula.MULTIPLE_AREAS)
        with mock.patch.object(table, 'MAX_VALUE_CHARACTERS', 3):
            self.assertEqual(self.read(selected()).code, formula.TOO_MUCH_TEXT)

    def test_iterator_failure_is_content_free(self):
        selection = selected()
        selection.Cells = []
        self.assertEqual(self.read(selection).code, formula.COM_FAILURE)

    def test_cancellation_before_during_after_read(self):
        for checks in ([True], [False,True], [False,False,True]):
            with self.subTest(checks=checks):
                self.assertEqual(self.read(selected(), cancelled=mock.Mock(side_effect=checks)).code, 'clipboard_changed')

    def test_selection_identity_change(self):
        identity = formula._selection_identity(selected())
        changed = tuple('Other' if i == 0 else value for i,value in enumerate(identity))
        self.assertEqual(self.read(selected(), identity=[identity, changed]).code, formula.SELECTION_CHANGED)

    def test_requires_two_equal_snapshots(self):
        one = self.read(selected())
        two = self.read(selected(display=('2.00','','',' x ')))
        for results, expected in (([one,two,two], formula.SUCCESS), ([one,two,one], formula.SELECTION_CHANGED), ([table.ExcelTableResult(formula.SELECTION_CHANGED),two,two], formula.SUCCESS)):
            with mock.patch.object(table, 'read_selected_excel_table', side_effect=results):
                self.assertEqual(table.read_stable_selected_excel_table().code, expected)

    def test_xml_fill_records_actual_source_cells(self):
        selection = ExcelTableSelection('Book.xlsx','Sheet1','$A$1:$C$3',3,3,
            ('그룹','항목','값','A','첫째','10',None,'둘째',None), ())
        off = ET.fromstring(table.selection_to_xml(selection, fill_blanks=False))
        on = ET.fromstring(table.selection_to_xml(selection, fill_blanks=True))
        self.assertEqual(off.get('빈칸채움수'), '0')
        self.assertEqual(on.get('빈칸채움수'), '1')
        self.assertEqual(on.get('빈칸채움셀'), '$A$3')
        self.assertEqual(on.get('주소'), '$A$1:$C$3')
        self.assertEqual(list(on)[-1][-1].text, None)


class GeneralXmlActionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        port.FormulaExportActionPortableTests.setUpClass.__func__(cls)
    _app = port.FormulaExportActionPortableTests._app

    def run_action(self, *, code=formula.SUCCESS, write_error=None):
        app = self._app()
        selection = ExcelTableSelection('Book.xlsx','Sheet1','$A$1:$B$2',2,2,('항목','값','한글','10'),())
        class ImmediateThread:
            def __init__(self, *, target, **kwargs): self.target=target
            def start(self): self.target()
        with mock.patch.object(self.module.threading,'Thread',ImmediateThread), mock.patch.object(self.module,'clipboard_change_count',return_value=41), mock.patch.object(self.module,'read_stable_selected_excel_table',return_value=table.ExcelTableResult(code,selection)) as reader, mock.patch.object(self.module,'write_text_only_clipboard',side_effect=write_error) as writer, mock.patch.object(self.module,'read_clipboard') as old_clipboard, mock.patch.object(self.module,'log'):
            app.copy_as_xml(None,None)
        self.assertFalse(app._formula_export_lock.locked())
        old_clipboard.assert_not_called()
        return app, writer, reader

    def test_current_selection_and_generation_write(self):
        app, writer, reader = self.run_action()
        self.assertEqual(ET.fromstring(writer.call_args.args[0]).get('주소'), '$A$1:$B$2')
        self.assertEqual(writer.call_args.kwargs['expected_change_count'],41)
        self.assertEqual(app._show_message_box_async.call_args.args[0], self.module.t('table_export.success','ko'))
        self.assertIn('is_cancelled',reader.call_args.kwargs)

    def test_read_failure_preserves_clipboard(self):
        app, writer, _ = self.run_action(code='display_overflow')
        writer.assert_not_called()
        self.assertEqual(app._show_message_box_async.call_args.args[0], self.module.t('table_export.error.display_overflow','ko'))

    def test_new_copy_and_failed_write_never_report_success(self):
        for error,key in ((self.module.ClipboardChangedError(),'clipboard_changed'),(self.module.ClipboardWriteError(),'clipboard_write_failed')):
            app, writer, _ = self.run_action(write_error=error)
            self.assertEqual(app._show_message_box_async.call_args.args[0], self.module.t('table_export.error.'+key,'ko'))

    def test_general_and_formula_share_gate_and_snapshot_setting(self):
        app=self._app(); app.fill_blanks=False
        workers=[]
        class HeldThread:
            def __init__(self, *, target, **kwargs): workers.append(target)
            def start(self): pass
        selection=ExcelTableSelection('Book.xlsx','Sheet1','$A$1:$B$2',2,2,('A','B','x','1'),())
        with mock.patch.object(self.module.threading,'Thread',HeldThread), mock.patch.object(self.module,'clipboard_change_count',return_value=41), mock.patch.object(self.module,'read_stable_selected_excel_table',return_value=table.ExcelTableResult(formula.SUCCESS,selection)), mock.patch.object(self.module,'selection_to_xml',return_value='<표/>') as serialize, mock.patch.object(self.module,'write_text_only_clipboard'), mock.patch.object(self.module,'read_clipboard') as read, mock.patch.object(self.module,'log'):
            app.copy_as_xml(None,None)
            app.copy_selected_excel_formulas(None,None)
            self.assertEqual(len(workers),1)
            self.assertFalse(app._augment_clipboard())
            read.assert_not_called()
            app.fill_blanks=True
            workers[0]()
            serialize.assert_called_once_with(selection,fill_blanks=False)
        self.assertFalse(app._formula_export_lock.locked())

if __name__ == '__main__': unittest.main()
