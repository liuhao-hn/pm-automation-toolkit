import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import core_logic as cl


class TestCalcRate(unittest.TestCase):
    def test_normal(self):
        self.assertEqual(cl.calc_rate(80, 100), 80.0)

    def test_zero_total(self):
        self.assertIsNone(cl.calc_rate(10, 0))


class TestCalcFluctuation(unittest.TestCase):
    def test_increase(self):
        self.assertEqual(cl.calc_fluctuation(110, 100), 10.0)

    def test_decrease(self):
        self.assertEqual(cl.calc_fluctuation(90, 100), -10.0)

    def test_none_and_zero(self):
        self.assertIsNone(cl.calc_fluctuation(None, 100))
        self.assertIsNone(cl.calc_fluctuation(100, 0))


class TestFlagWarn(unittest.TestCase):
    def test_excellent(self):
        self.assertIn("优秀", cl.flag_warn(15.0))

    def test_abnormal(self):
        self.assertIn("异常", cl.flag_warn(-12.0))

    def test_flat(self):
        self.assertEqual(cl.flag_warn(5.0), "")

    def test_none(self):
        self.assertEqual(cl.flag_warn(None), "")


class TestFlagCombined(unittest.TestCase):
    def test_abnormal_negative(self):
        self.assertIn("异常", cl.flag_combined(-15.0, 95.0, 90.0))

    def test_excellent(self):
        self.assertIn("优秀", cl.flag_combined(15.0, 95.0, 90.0))

    def test_qualified(self):
        self.assertIn("合格", cl.flag_combined(5.0, 95.0, 90.0))

    def test_unqualified(self):
        self.assertIn("不合格", cl.flag_combined(5.0, 80.0, 90.0))

    def test_no_data(self):
        self.assertEqual(cl.flag_combined(None, None, 90.0), "")

    def test_nan_no_data(self):
        self.assertEqual(cl.flag_combined(float("nan"), float("nan"), 90.0), "")


class TestValid(unittest.TestCase):
    def test_valid(self):
        self.assertTrue(cl._valid(5.0))

    def test_none(self):
        self.assertFalse(cl._valid(None))

    def test_nan(self):
        self.assertFalse(cl._valid(float("nan")))


if __name__ == "__main__":
    unittest.main()
