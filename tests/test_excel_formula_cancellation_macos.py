from __future__ import annotations

import threading
import unittest
from unittest.mock import patch

from Foundation import NSAppleScript

from tablemark import excel_formula as ef
from tablemark.converter.formula_export import formula_selection_to_ai_xml
from tests.test_excel_formula_macos import (
    FakeExecutor, area_payload, mask_payload, reference_area, reference_payload,
    value_area,
)


class ScriptedExecutor(FakeExecutor):
    def __init__(self, payloads):
        super().__init__(payloads)
        self.thread_ids = []

    def run(self, source):
        self.sources.append(source)
        self.thread_ids.append(threading.get_ident())
        if not self.payloads:
            raise AssertionError("unexpected executor call")
        value = self.payloads.pop(0)
        if isinstance(value, Exception):
            raise value
        return value


def fixture(target_count=33, *, reverse_owner=False, expression=None):
    refs = [f"B{row}" for row in range(1, target_count + 1)]
    expression = expression or "=" + "+".join(refs)
    rows = 2 if reverse_owner else 1
    address = "$A$1:$A$2" if reverse_owner else "$A$1"
    mask = mask_payload(address=address, rows=rows, mask=["true"] * rows)
    expressions = [expression]
    if reverse_owner:
        expressions.append("=" + "+".join(reversed(refs)))
    areas = [[address, str(rows), "1", [[v] for v in expressions], "present", [[v] for v in expressions]]]
    values = value_area(["0"] * rows, address=address, rows=rows)
    payload = area_payload(areas, address=address, value_areas=[values])
    plan = ef._parse_mask_result(mask)
    selection = ef.parse_excel_formula_result(payload, expected_plan=plan)
    requests, completeness, issues = ef._reference_read_plan(selection)
    raw = [reference_area(r.owner_address, r.target.sheet, r.target.address,
                          [f"v{index + 1}"], rows=r.target.row_count,
                          columns=r.target.column_count)
           for index, r in enumerate(requests)]
    return mask, payload, plan, selection, requests, completeness, issues, raw


def batch_payloads(data, *, value_prefix=None):
    _mask, _payload, _plan, selection, requests, _complete, _issues, raw = data
    copied = []
    for index, item in enumerate(raw):
        item = list(item)
        if value_prefix is not None:
            item[-1] = [["value", f"{value_prefix}{index + 1}"]]
        copied.append(item)
    result = []
    count = 0
    start = 0
    batches = ef._reference_batches(requests)
    for batch in batches:
        values = copied[start:start + len(batch)]
        count += sum(len(x[1]) for area in values for x in area[-1])
        payload = reference_payload(values, address=selection.address)
        if len(batches) > 1:
            payload.append(str(count))
        result.append(payload)
        start += len(batch)
    return result


class FormulaCancellationTests(unittest.TestCase):
    def test_cancellation_codes_remain_distinguishable(self):
        for code in (ef.CANCELLED, ef.CLIPBOARD_CHANGED):
            self.assertEqual(ef.ExcelFormulaError(code).code, code)

    def test_cancel_at_each_native_phase_stops_without_partial_result_or_retry(self):
        data = fixture(33)
        for code in (ef.CANCELLED, ef.CLIPBOARD_CHANGED):
            for completed_calls in (0, 1, 2, 3, 4, 5):
                with self.subTest(code=code, completed_calls=completed_calls):
                    executor = ScriptedExecutor([data[0], data[1], *batch_payloads(data)])

                    def check():
                        if len(executor.sources) >= completed_calls:
                            raise ef.ExcelFormulaError(code)

                    with self.assertRaises(ef.ExcelFormulaError) as caught:
                        ef.read_stable_selected_excel_formulas(executor, check_cancelled=check)
                    self.assertEqual(caught.exception.code, code)
                    self.assertEqual(len(executor.sources), completed_calls)

    def test_cancel_between_completed_snapshots_skips_the_second_snapshot(self):
        data = fixture(17)
        executor = ScriptedExecutor([data[0], data[1], *batch_payloads(data)])
        cancelled = False
        original = ef._parse_reference_result

        def parse(*args, **kwargs):
            nonlocal cancelled
            result = original(*args, **kwargs)
            cancelled = True
            return result

        def check():
            if cancelled:
                raise ef.ExcelFormulaError(ef.CLIPBOARD_CHANGED)

        with patch.object(ef, "_parse_reference_result", side_effect=parse):
            with self.assertRaises(ef.ExcelFormulaError) as caught:
                ef.read_stable_selected_excel_formulas(executor, check_cancelled=check)
        self.assertEqual(caught.exception.code, ef.CLIPBOARD_CHANGED)
        self.assertEqual(len(executor.sources), 4)

    def test_callback_failure_during_references_is_not_swallowed_as_partial(self):
        data = fixture(33)
        executor = ScriptedExecutor([data[0], data[1], *batch_payloads(data)])

        def check():
            if len(executor.sources) == 3:
                raise RuntimeError("PRIVATE_WORKBOOK_CONTENT")

        with self.assertRaises(ef.ExcelFormulaError) as caught:
            ef.read_selected_excel_formulas(executor, check_cancelled=check)
        self.assertEqual(caught.exception.code, ef.EXECUTION_FAILED)
        self.assertNotIn("PRIVATE", str(caught.exception))
        self.assertEqual(len(executor.sources), 3)

    def test_cancellation_wins_if_native_call_also_fails(self):
        data = fixture(17)
        executor = ScriptedExecutor([data[0], data[1], RuntimeError("native failure")])

        def check():
            if len(executor.sources) >= 3:
                raise ef.ExcelFormulaError(ef.CANCELLED)

        with self.assertRaises(ef.ExcelFormulaError) as caught:
            ef.read_selected_excel_formulas(executor, check_cancelled=check)
        self.assertEqual(caught.exception.code, ef.CANCELLED)
        self.assertEqual(len(executor.sources), 3)

    def test_reference_error_payload_cancellation_is_never_degraded(self):
        data = fixture(1)
        for code in (ef.CANCELLED, ef.CLIPBOARD_CHANGED):
            with self.subTest(code=code):
                executor = ScriptedExecutor([data[0], data[1], ["error", code]])
                with self.assertRaises(ef.ExcelFormulaError) as caught:
                    ef.read_stable_selected_excel_formulas(executor)
                self.assertEqual(caught.exception.code, code)
                self.assertEqual(len(executor.sources), 3)

    def test_no_callback_preserves_small_export_and_caller_thread(self):
        data = fixture(2)
        executor = ScriptedExecutor([data[0], data[1], *batch_payloads(data)])
        expected = ef._parse_reference_result(reference_payload(data[-1]), data[3],
                                              data[4], dict(data[5]), data[6])
        actual = ef.read_selected_excel_formulas(executor)
        self.assertEqual(actual, expected)
        self.assertEqual(len(executor.sources), 3)
        self.assertEqual(set(executor.thread_ids), {threading.get_ident()})

    def test_without_references_cancel_after_values_never_returns_snapshot(self):
        data = fixture(expression="=1+1")
        executor = ScriptedExecutor([data[0], data[1]])

        def check():
            if len(executor.sources) >= 2:
                raise ef.ExcelFormulaError(ef.CLIPBOARD_CHANGED)

        with self.assertRaises(ef.ExcelFormulaError) as caught:
            ef.read_selected_excel_formulas(executor, check_cancelled=check)
        self.assertEqual(caught.exception.code, ef.CLIPBOARD_CHANGED)
        self.assertEqual(len(executor.sources), 2)


class FormulaReferenceBatchTests(unittest.TestCase):
    def test_batch_range_bound_keeps_exact_target_order(self):
        data = fixture(33)
        batches = ef._reference_batches(data[4])
        self.assertEqual([len(b) for b in batches], [16, 16, 1])
        self.assertEqual([r.target.address for b in batches for r in b],
                         [f"$B${row}" for row in range(1, 34)])

    def test_batch_cell_bound_does_not_split_or_expand_ranges(self):
        data = fixture(expression="=SUM(B1:B1025)+SUM(C1:C1025)+SUM(D1:D2048)")
        batches = ef._reference_batches(data[4])
        self.assertEqual([len(b) for b in batches], [1, 1, 1])
        self.assertEqual([b[0].target.address for b in batches],
                         ["$B$1:$B$1025", "$C$1:$C$1025", "$D$1:$D$2048"])
        self.assertTrue(all(sum(r.target.row_count*r.target.column_count for r in b) <= 2048
                            for b in batches))

    def test_global_range_limit_applies_once_before_batches(self):
        data = fixture(260)
        self.assertEqual(len(data[4]), 256)
        self.assertEqual(sum(map(len, ef._reference_batches(data[4]))), 256)
        self.assertFalse(data[5]["$A$1"])
        self.assertEqual(data[6]["$A$1"], ("range_count_limit",))

    def test_global_cell_limit_applies_once_before_batches(self):
        expression = "=" + "+".join(f"SUM({col}1:{col}2000)" for col in "BCDEFG")
        data = fixture(expression=expression)
        self.assertEqual(len(data[4]), 5)
        self.assertEqual(sum(r.target.row_count*r.target.column_count for r in data[4]), 10_000)
        self.assertEqual(len(ef._reference_batches(data[4])), 5)
        self.assertEqual(data[6]["$A$1"], ("cell_count_limit",))

    def test_batched_output_matches_original_payload_and_owner_order(self):
        data = fixture(33, reverse_owner=True)
        expected = ef._parse_reference_result(
            reference_payload(data[-1], address=data[3].address), data[3], data[4],
            dict(data[5]), data[6])
        actual = ef.read_selected_excel_formulas(
            ScriptedExecutor([data[0], data[1], *batch_payloads(data)]))
        self.assertEqual(actual, expected)
        self.assertEqual(formula_selection_to_ai_xml(actual), formula_selection_to_ai_xml(expected))
        self.assertEqual([r.address for r in actual.cells[1].references],
                         [f"$B${row}" for row in range(33, 0, -1)])
        self.assertIs(actual.cells[0].references[0], actual.cells[1].references[-1])

    def test_native_character_budget_is_carried_across_batches(self):
        data = fixture(3)
        first = reference_area("$A$1", "Sheet1", "$B$1", ["123456"])
        last = reference_area("$A$1", "Sheet1", "$B$3", ["123"])
        responses = [reference_payload([first])+["6"],
                     reference_payload([["too_large", "$A$1", "Sheet1", "$B$2"]])+["6"],
                     reference_payload([last])+["9"]]
        executor = ScriptedExecutor([data[0], data[1], *responses])
        with patch.object(ef, "MAX_REFERENCE_BATCH_RANGES", 1), patch.object(ef, "MAX_VALUE_CHARACTERS", 10):
            result = ef.read_selected_excel_formulas(executor)
        self.assertIn("set referenceValueCharacters to 0", executor.sources[2])
        self.assertIn("set referenceValueCharacters to 6", executor.sources[3])
        self.assertIn("set referenceValueCharacters to 6", executor.sources[4])
        self.assertEqual([r.address for r in result.cells[0].references], ["$B$1", "$B$3"])
        self.assertEqual(result.cells[0].reference_issues, ("value_size_limit",))

    def test_python_character_limit_is_global_including_selected_values(self):
        data = fixture(2)
        values = [reference_area("$A$1", "Sheet1", "$B$1", ["123456"]),
                  reference_area("$A$1", "Sheet1", "$B$2", ["1234"])]
        executor = ScriptedExecutor([data[0], data[1], reference_payload(values[:1])+["6"],
                                     reference_payload(values[1:])+["10"]])
        with patch.object(ef, "MAX_REFERENCE_BATCH_RANGES", 1), patch.object(ef, "MAX_VALUE_CHARACTERS", 10):
            result = ef.read_selected_excel_formulas(executor)
        self.assertEqual([r.address for r in result.cells[0].references], ["$B$1"])
        self.assertEqual(result.cells[0].reference_issues, ("value_size_limit",))

    def test_invalid_cumulative_count_discards_all_reference_batches(self):
        data = fixture(17)
        for bad_count in ("0", str(ef.MAX_VALUE_CHARACTERS+1), "-1", True):
            with self.subTest(bad_count=bad_count):
                responses = batch_payloads(data)
                responses[1][-1] = bad_count
                result = ef.read_selected_excel_formulas(
                    ScriptedExecutor([data[0], data[1], *responses]))
                self.assertEqual(result.cells[0].references, ())
                self.assertEqual(result.cells[0].reference_issues, ("read_failed",))

    def test_middle_native_failure_keeps_established_all_reference_degradation(self):
        data = fixture(33)
        responses = batch_payloads(data)
        executor = ScriptedExecutor([data[0], data[1], responses[0], RuntimeError("PRIVATE")])
        result = ef.read_selected_excel_formulas(executor)
        self.assertEqual(result.cells[0].references, ())
        self.assertEqual(result.cells[0].reference_issues, ("read_failed",))
        self.assertEqual(result.cells[0].formula_a1, data[3].cells[0].formula_a1)
        self.assertEqual(result.cells[0].value, data[3].cells[0].value)
        self.assertEqual(len(executor.sources), 4)

    def test_per_target_read_failure_keeps_other_targets(self):
        data = fixture(17)
        responses = batch_payloads(data)
        responses[0][4][0] = ["unresolved", "$A$1", "Sheet1", "$B$1"]
        result = ef.read_selected_excel_formulas(ScriptedExecutor([data[0], data[1], *responses]))
        self.assertEqual(len(result.cells[0].references), 16)
        self.assertEqual(result.cells[0].reference_issues, ("read_failed",))

    def test_batch_identity_mismatch_discards_snapshot_and_retries_fresh(self):
        data = fixture(17)
        bad = batch_payloads(data, value_prefix="old")
        bad[1][1] = "Other.xlsx"
        fresh = batch_payloads(data, value_prefix="new")
        executor = ScriptedExecutor([data[0], data[1], *bad,
                                     data[0], data[1], *fresh,
                                     data[0], data[1], *batch_payloads(data, value_prefix="new")])
        result = ef.read_stable_selected_excel_formulas(executor)
        self.assertTrue(all(r.cells[0].value.startswith("new") for r in result.cells[0].references))
        self.assertEqual(len(executor.sources), 12)

    def test_failure_between_equal_snapshots_breaks_consecutiveness(self):
        data = fixture(17)
        bad = batch_payloads(data)
        bad[1][3] = "$A$2"
        executor = ScriptedExecutor([data[0], data[1], *batch_payloads(data),
                                     data[0], data[1], *bad,
                                     data[0], data[1], *batch_payloads(data)])
        with self.assertRaises(ef.ExcelFormulaError) as caught:
            ef.read_stable_selected_excel_formulas(executor)
        self.assertEqual(caught.exception.code, ef.SELECTION_CHANGED)
        self.assertEqual(len(executor.sources), 12)

    def test_batch_scripts_compile_and_retain_pre_post_selection_identity(self):
        data = fixture(17)
        for initial in (0, 15):
            source = ef._build_reference_read_script(
                data[2], ef._reference_batches(data[4])[0],
                initial_reference_characters=initial, include_character_count=True)
            script = NSAppleScript.alloc().initWithSource_(source)
            compiled, error = script.compileAndReturnError_(None)
            self.assertTrue(compiled, error)
            self.assertIsNone(error)
            self.assertIn("get address selectedRange", source)
            self.assertIn("get address selection", source)
            self.assertIn("referenceValueCharacters as text", source)
            self.assertNotIn("activate", source)


def selection_chunk_fixture(rows=501, columns=8, *, dense=False, version="v"):
    address = ef._absolute_address(1, 1, rows, columns)
    flags = [
        dense or (row > 0 and column == min(5, columns-1))
        for row in range(rows) for column in range(columns)
    ]
    mask = mask_payload(address=address, rows=rows, columns=columns,
                        mask="".join("1" if flag else "0" for flag in flags))
    plan = ef._parse_mask_result(mask)
    values = [
        [chunk.address, str(chunk.row_count), str(chunk.column_count),
         [["text", f"{version}{row}:{column}"]
          for row in range(chunk.start_row, chunk.end_row+1)
          for column in range(chunk.start_column, chunk.end_column+1)]]
        for chunk in plan.value_chunks
    ]
    formulas = [
        [chunk.address, str(chunk.row_count), str(chunk.column_count),
         [["=1+1"] * chunk.column_count for _ in range(chunk.row_count)],
         "present",
         [["=1+1"] * chunk.column_count for _ in range(chunk.row_count)]]
        for chunk in plan.chunks
    ]
    full = area_payload(formulas, address=address, value_areas=values)
    return mask, plan, full


def selection_chunk_payloads(data):
    _mask, plan, full = data
    value_by_address = {area[0]: area for area in full[6]}
    formula_by_address = {area[0]: area for area in full[7]}
    values_count = formulas_count = 0
    payloads = []
    batches = ef._formula_read_batches(plan)
    for value_chunks, formula_chunks in batches:
        values = [value_by_address[c.address] for c in value_chunks]
        formulas = [formula_by_address[c.address] for c in formula_chunks]
        values_count += sum(len(value[1]) for area in values for value in area[-1])
        for area in formulas:
            formulas_count += sum(len(value) for row in area[3] for value in row)
            if area[4] == "present":
                formulas_count += sum(len(value) for row in area[5] for value in row)
        payload = [*full[:6], values, formulas]
        if len(batches) > 1:
            payload.extend([str(values_count), str(formulas_count)])
        payloads.append(payload)
    return payloads


class FormulaValueBatchTests(unittest.TestCase):
    def test_large_selection_chunk_sequence_and_output_match_previous_payload(self):
        data = selection_chunk_fixture()
        plan = data[1]
        batches = ef._formula_read_batches(plan)
        self.assertEqual(
            [(sum(c.row_count*c.column_count for c in values),
              sum(c.row_count*c.column_count for c in formulas))
             for values, formulas in batches],
            [(2048, 0), (1960, 0), (0, 500)],
        )
        self.assertEqual([c for values, _ in batches for c in values], list(plan.value_chunks))
        self.assertEqual([c for _, formulas in batches for c in formulas], list(plan.chunks))
        executor = ScriptedExecutor([data[0], *selection_chunk_payloads(data)])
        expected = ef.parse_excel_formula_result(data[2], expected_plan=plan)
        actual = ef.read_selected_excel_formulas(executor)
        self.assertEqual(actual, expected)
        self.assertEqual(formula_selection_to_ai_xml(actual), formula_selection_to_ai_xml(expected))
        self.assertEqual(len(executor.sources), 4)
        self.assertEqual(set(executor.thread_ids), {threading.get_ident()})

    def test_small_single_call_uses_total_value_and_two_formula_array_work(self):
        for rows, expected_batches in ((682, 1), (683, 2), (2048, 2)):
            with self.subTest(rows=rows):
                data = selection_chunk_fixture(rows, 1, dense=True)
                self.assertEqual(len(data[1].value_chunks), 1)
                batches = ef._formula_read_batches(data[1])
                self.assertEqual(len(batches), expected_batches)
                payloads = selection_chunk_payloads(data)
                self.assertEqual(len(payloads[0]), 8 if expected_batches == 1 else 10)
                actual = ef.read_selected_excel_formulas(ScriptedExecutor([data[0], *payloads]))
                self.assertEqual(actual, ef.parse_excel_formula_result(data[2], data[1]))

    def test_sparse_chunk_count_also_bounds_small_selection_native_calls(self):
        data = fixture()
        mask = mask_payload(address="$A$1:$P$8", rows=8, columns=16,
                            mask="".join("1" if (r+c)%2 == 0 else "0"
                                         for r in range(8) for c in range(16)))
        plan = ef._parse_mask_result(mask)
        self.assertEqual(len(plan.chunks), 64)
        batches = ef._formula_read_batches(plan)
        self.assertEqual(len(batches), 65)
        self.assertEqual([c for _, formulas in batches for c in formulas], list(plan.chunks))
        self.assertTrue(all(len(values)+len(formulas) == 1 for values, formulas in batches))

    def test_cancellation_at_every_value_and_formula_chunk_stops_before_next_call(self):
        data = selection_chunk_fixture(100, 100, dense=True)
        self.assertEqual(len(ef._formula_read_batches(data[1])), 10)
        for code in (ef.CANCELLED, ef.CLIPBOARD_CHANGED):
            for completed in range(12):
                with self.subTest(code=code, completed=completed):
                    executor = ScriptedExecutor([data[0], *selection_chunk_payloads(data)])
                    def check():
                        if len(executor.sources) >= completed:
                            raise ef.ExcelFormulaError(code)
                    with self.assertRaises(ef.ExcelFormulaError) as caught:
                        ef.read_stable_selected_excel_formulas(executor, check_cancelled=check)
                    self.assertEqual(caught.exception.code, code)
                    self.assertEqual(len(executor.sources), completed)

    def test_value_and_formula_native_character_counters_are_cumulative(self):
        with patch.object(ef, "MAX_FORMULA_CHUNK_CELLS", 2), patch.object(
            ef, "MAX_FORMULA_READ_BATCH_WORK_CELLS", 1
        ):
            data = selection_chunk_fixture(4, 1, dense=True)
            payloads = selection_chunk_payloads(data)
            executor = ScriptedExecutor([data[0], *payloads])
            ef.read_selected_excel_formulas(executor)
        first_values, all_values = payloads[0][8], payloads[1][8]
        first_formulas = payloads[2][9]
        self.assertIn("set valueCharacterCount to 0", executor.sources[1])
        self.assertIn(f"set valueCharacterCount to {first_values}", executor.sources[2])
        self.assertIn(f"set valueCharacterCount to {all_values}", executor.sources[3])
        self.assertIn(f"set valueCharacterCount to {all_values}", executor.sources[4])
        self.assertIn("set formulaCharacterCount to 0", executor.sources[3])
        self.assertIn(f"set formulaCharacterCount to {first_formulas}", executor.sources[4])

    def test_full_parser_enforces_global_character_limits_even_if_counts_underreport(self):
        with patch.object(ef, "MAX_FORMULA_CHUNK_CELLS", 2), patch.object(
            ef, "MAX_FORMULA_READ_BATCH_WORK_CELLS", 1
        ):
            data = selection_chunk_fixture(4, 1, dense=True)
            for field, limit in (("MAX_VALUE_CHARACTERS", 12), ("MAX_FORMULA_CHARACTERS", 20)):
                with self.subTest(field=field):
                    payloads = selection_chunk_payloads(data)
                    for payload in payloads:
                        payload[8:] = ["0", "0"]
                    executor = ScriptedExecutor([data[0], *payloads])
                    with patch.object(ef, field, limit):
                        with self.assertRaises(ef.ExcelFormulaError) as caught:
                            ef.read_selected_excel_formulas(executor)
                    self.assertEqual(caught.exception.code, ef.TOO_MUCH_TEXT)
                    self.assertEqual(len(executor.sources), 5)

    def test_bad_value_or_formula_counter_fails_before_the_next_chunk(self):
        data = selection_chunk_fixture()
        for field, limit in ((8, ef.MAX_VALUE_CHARACTERS), (9, ef.MAX_FORMULA_CHARACTERS)):
            for invalid in (True, "-1", str(limit+1)):
                with self.subTest(field=field, invalid=invalid):
                    payloads = selection_chunk_payloads(data)
                    payloads[0][field] = invalid
                    executor = ScriptedExecutor([data[0], *payloads])
                    with self.assertRaises(ef.ExcelFormulaError) as caught:
                        ef.read_selected_excel_formulas(executor)
                    self.assertEqual(caught.exception.code, ef.INVALID_RESPONSE)
                    self.assertEqual(len(executor.sources), 2)
        payloads = selection_chunk_payloads(data)
        payloads[1][8] = "0"
        executor = ScriptedExecutor([data[0], *payloads])
        with self.assertRaises(ef.ExcelFormulaError) as caught:
            ef.read_selected_excel_formulas(executor)
        self.assertEqual(caught.exception.code, ef.INVALID_RESPONSE)
        self.assertEqual(len(executor.sources), 3)

    def test_missing_extra_wrong_chunk_and_cross_phase_area_fail_closed(self):
        data = selection_chunk_fixture()
        for change in ("missing_value", "extra_value", "wrong_value", "formula_during_value",
                       "missing_formula", "wrong_formula"):
            with self.subTest(change=change):
                payloads = selection_chunk_payloads(data)
                # Own the area lists so source fixtures are never mutated.
                payloads = [[*p[:6], list(p[6]), list(p[7]), *p[8:]] for p in payloads]
                expected_calls = 2
                if change == "missing_value":
                    payloads[0][6] = []
                elif change == "extra_value":
                    payloads[0][6].append(payloads[0][6][0])
                elif change == "wrong_value":
                    payloads[0][6] = list(payloads[1][6])
                elif change == "formula_during_value":
                    payloads[0][7] = list(payloads[-1][7])
                elif change == "missing_formula":
                    payloads[-1][7] = []
                    expected_calls = 4
                else:
                    area = list(payloads[-1][7][0])
                    area[0] = "$G$2:$G$501"
                    payloads[-1][7] = [area]
                    expected_calls = 4
                executor = ScriptedExecutor([data[0], *payloads])
                with self.assertRaises(ef.ExcelFormulaError) as caught:
                    ef.read_selected_excel_formulas(executor)
                self.assertEqual(caught.exception.code, ef.INVALID_RESPONSE)
                self.assertEqual(len(executor.sources), expected_calls)

    def test_middle_chunk_native_failure_or_text_limit_never_degrades_to_success(self):
        data = selection_chunk_fixture()
        for failure, code in ((RuntimeError("PRIVATE"), ef.EXECUTION_FAILED),
                              (["error", ef.TOO_MUCH_TEXT], ef.TOO_MUCH_TEXT)):
            with self.subTest(code=code):
                payloads = selection_chunk_payloads(data)
                executor = ScriptedExecutor([data[0], payloads[0], failure, *payloads[2:]])
                with self.assertRaises(ef.ExcelFormulaError) as caught:
                    ef.read_stable_selected_excel_formulas(executor)
                self.assertEqual(caught.exception.code, code)
                self.assertNotIn("PRIVATE", str(caught.exception))
                self.assertEqual(len(executor.sources), 3)

    def test_middle_identity_change_retries_two_complete_fresh_snapshots(self):
        before = selection_chunk_fixture(version="old")
        after = selection_chunk_fixture(version="new")
        bad = selection_chunk_payloads(before)
        bad[1][2] = "Other sheet"
        executor = ScriptedExecutor([
            before[0], *bad[:2],
            after[0], *selection_chunk_payloads(after),
            after[0], *selection_chunk_payloads(after),
        ])
        actual = ef.read_stable_selected_excel_formulas(executor)
        self.assertEqual(actual, ef.parse_excel_formula_result(after[2], after[1]))
        self.assertEqual(len(executor.sources), 11)

    def test_partial_batch_between_equal_snapshots_breaks_consecutiveness(self):
        data = selection_chunk_fixture()
        bad = selection_chunk_payloads(data)
        bad[1][3] = "$A$2:$H$502"
        executor = ScriptedExecutor([
            data[0], *selection_chunk_payloads(data),
            data[0], *bad[:2],
            data[0], *selection_chunk_payloads(data),
        ])
        with self.assertRaises(ef.ExcelFormulaError) as caught:
            ef.read_stable_selected_excel_formulas(executor)
        self.assertEqual(caught.exception.code, ef.SELECTION_CHANGED)
        self.assertEqual(len(executor.sources), 11)

    def test_completed_snapshots_with_same_selection_and_changed_values_need_a_third_read(self):
        before = selection_chunk_fixture(version="old")
        after = selection_chunk_fixture(version="new")
        executor = ScriptedExecutor([
            before[0], *selection_chunk_payloads(before),
            after[0], *selection_chunk_payloads(after),
            after[0], *selection_chunk_payloads(after),
        ])
        actual = ef.read_stable_selected_excel_formulas(executor)
        self.assertEqual(actual, ef.parse_excel_formula_result(after[2], after[1]))
        self.assertEqual(len(executor.sources), 12)

    def test_missing_r1c1_retains_existing_fallback_without_missing_a1_ranges(self):
        data = selection_chunk_fixture()
        data[2][7][0][4:] = ["missing", ""]
        executor = ScriptedExecutor([data[0], *selection_chunk_payloads(data)])
        actual = ef.read_selected_excel_formulas(executor)
        self.assertEqual(actual, ef.parse_excel_formula_result(data[2], data[1]))
        self.assertTrue(all(cell.formula_r1c1 is None for cell in actual.cells))

    def test_each_generated_batch_compiles_and_checks_full_identity_and_topology(self):
        data = selection_chunk_fixture()
        for values, formulas in ef._formula_read_batches(data[1]):
            source = ef._build_formula_read_script(
                data[1], value_chunks=values, formula_chunks=formulas,
                initial_value_characters=17, initial_formula_characters=23,
                include_character_counts=True,
            )
            script = NSAppleScript.alloc().initWithSource_(source)
            compiled, error = script.compileAndReturnError_(None)
            self.assertTrue(compiled, error)
            self.assertIsNone(error)
            self.assertIn("set expectedFormulaCount to 500", source)
            self.assertIn('set expectedSelectionAddress to "$A$1:$H$501"', source)
            self.assertIn("name of active workbook", source)
            self.assertIn("name of active sheet", source)
            self.assertIn("get address currentRange", source)
            self.assertIn("formulaTopologyMatches(selectedRange", source)
            self.assertIn("formulaTopologyMatches(currentRange", source)
            self.assertIn("set valueCharacterCount to 17", source)
            self.assertIn("set formulaCharacterCount to 23", source)
            self.assertIn("valueCharacterCount as text, formulaCharacterCount as text", source)
            self.assertNotIn("activate", source)


if __name__ == "__main__":
    unittest.main()
