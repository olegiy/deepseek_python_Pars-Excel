# tests/test_core_functions.py

import unittest
from data_processing import try_convert, extract_thickness_value
from pricing import calculate_prices_for_section
from openpyxl import Workbook

class TestCoreFunctions(unittest.TestCase):

    def test_try_convert_number(self):
        self.assertEqual(try_convert(12.3), (13, '0'))
        self.assertEqual(try_convert(5), (5, '0'))

    def test_try_convert_string_with_comma(self):
        self.assertEqual(try_convert("12,5"), (13, '0'))

    def test_try_convert_fraction(self):
        self.assertEqual(try_convert("1/2"), (1, '0'))

    def test_try_convert_invalid(self):
        self.assertEqual(try_convert("abc"), ("abc", None))
        self.assertEqual(try_convert(None), (None, None))

    def test_extract_thickness_value(self):
        self.assertEqual(extract_thickness_value("Толщина стенки: 5,5 мм", "толщина стенки"), 5.5)
        self.assertEqual(extract_thickness_value("Средняя толщина ноги: 10 мм", "средняя толщина ноги"), 10.0)
        self.assertIsNone(extract_thickness_value("Нет данных", "толщина"))

    def test_calculate_prices_for_section(self):
        # Создаем тестовый workbook и листы
        wb = Workbook()
        ws = wb.active
        ws.title = "Part Info"

        price_wb = Workbook()
        price_ws = price_wb.active
        price_ws.title = "Price Data"
        price_ws.append(["Thickness", "Dummy", "Contour Price", "Cut Price"])
        price_ws.append([5, "", 100, 20])

        # Подготовка тестовых данных
        ws.append(["Section: Test Thickness: 5"])
        ws.append(["ID", "Qty", "Part Length(mm)", "Contour Qty", "Cut Length(mm)", "Price(₽)"])
        ws.append([1, 2, 1000, 1, 2000, None])  # ID 1, Qty 2, Part Length 1m, Contour 1, Cut 2m

        section_row = 1
        next_section_row = 3

        total_price = calculate_prices_for_section(ws, price_ws, section_row, next_section_row)

        # Проверяем цену: (1 * 100) + (2 / 1000 * 20) = 100 + 0.04 = ~100.04 за единицу * 2 шт = 200.08
        self.assertAlmostEqual(total_price, 200.08, delta=0.01)


if __name__ == '__main__':
    unittest.main()