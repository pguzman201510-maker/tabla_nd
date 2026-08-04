#!/usr/bin/env python3
"""
Unit tests/verification for the DGCP PDF parsing logic.
"""

import unittest
from process_dgcp_pdfs import parse_table_row

class TestDGCPParsing(unittest.TestCase):

    def test_standard_row_with_ref_and_date(self):
        line = "2013-06-15 0.00 0.00 0.00 9,383.56 0.00 1302230-01 2013-06-17"
        parsed = parse_table_row(line)
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed['Venc_Izq'], "2013-06-15")
        self.assertEqual(parsed['Desembolsos'], 0.0)
        self.assertEqual(parsed['Amortizaciones'], 0.0)
        self.assertEqual(parsed['Intereses'], 0.0)
        self.assertEqual(parsed['Comisiones'], 9383.56)
        self.assertEqual(parsed['Saldo_Pactado'], 0.0)
        self.assertEqual(parsed['Ref'], "1302230-01")
        self.assertEqual(parsed['Fecha_Der'], "2013-06-17")

    def test_row_with_non_zero_amortization_and_ref(self):
        line = "2044-02-26 0.00 15,997,000.00 0.00 0.00 2,484,003,000.00 2600015-06 2026-04-08"
        parsed = parse_table_row(line)
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed['Venc_Izq'], "2044-02-26")
        self.assertEqual(parsed['Desembolsos'], 0.0)
        self.assertEqual(parsed['Amortizaciones'], 15997000.0)
        self.assertEqual(parsed['Intereses'], 0.0)
        self.assertEqual(parsed['Comisiones'], 0.0)
        self.assertEqual(parsed['Saldo_Pactado'], 2484003000.0)
        self.assertEqual(parsed['Ref'], "2600015-06")
        self.assertEqual(parsed['Fecha_Der'], "2026-04-08")

    def test_row_with_programada_marker_no_ref(self):
        line = "2044-02-26 0.00 2,019,978,000.00P 56,811,881.25 P 0.00 0.00"
        parsed = parse_table_row(line)
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed['Venc_Izq'], "2044-02-26")
        self.assertEqual(parsed['Desembolsos'], 0.0)
        self.assertEqual(parsed['Amortizaciones'], 2019978000.0)
        self.assertEqual(parsed['Intereses'], 56811881.25)
        self.assertEqual(parsed['Comisiones'], 0.0)
        self.assertEqual(parsed['Saldo_Pactado'], 0.0)
        self.assertEqual(parsed['Ref'], "")
        self.assertEqual(parsed['Fecha_Der'], "2044-02-26")

    def test_row_with_programada_marker_and_p_suffix(self):
        line = "2030-07-15 0.00 17,485,934.77P 426,697.85 P 0.00 0.00"
        parsed = parse_table_row(line)
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed['Venc_Izq'], "2030-07-15")
        self.assertEqual(parsed['Desembolsos'], 0.0)
        self.assertEqual(parsed['Amortizaciones'], 17485934.77)
        self.assertEqual(parsed['Intereses'], 426697.85)
        self.assertEqual(parsed['Comisiones'], 0.0)
        self.assertEqual(parsed['Saldo_Pactado'], 0.0)
        self.assertEqual(parsed['Ref'], "")
        self.assertEqual(parsed['Fecha_Der'], "2030-07-15")

if __name__ == "__main__":
    unittest.main()
