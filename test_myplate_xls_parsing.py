import unittest
import os
import warnings
from datetime import datetime
import xlwt
import pandas as pd
import numpy as np
from helpers import formatted_datestring, write_sheet_data
from mypparser import MyPlateDetailedExportParser as mpp

XLS_TEST_FILENAME = "example.xls"


def testfile_cleanup_helper(path):
    if os.path.isfile(path):
        os.remove(path)


class TestBasic(unittest.TestCase):
    def setUp(self):
        warnings.filterwarnings(action="ignore", category=FutureWarning)

        testfile_cleanup_helper(XLS_TEST_FILENAME)

        # Load test data
        self.first_day_date = datetime(2023, 5, 25)
        self.second_day_date = datetime(2023, 5, 26)
        self.first_day_datestring = formatted_datestring(self.first_day_date)
        self.second_day_datestring = formatted_datestring(self.second_day_date)

        self.first_day_rows = [
            ["Date:", self.first_day_datestring],
            [
                "Meal",
                "Brand",
                "Name",
                "Servings",
                "Calories",
                "Nutrient A",
                "Nutrient B",
                "Nutrient C",
            ],
            ["breakfast", None, "eggs", 2, 180, 19, "12mg", 89],
            ["lunch", None, "turkey sandwich", 0.75, 488, 38, 22.8, 75],
            ["snacks", None, "apple", 1, 90, 2, 0, 12],
            ["dinner", None, "chicken burrito", 1, 850, 47, 32, 22],
        ]

        filler_rows = [
            [None],
            [None],
            ["Junk", "should", "not", "be", "included", "in", "final", 123, None],
            [None],
            [None],
        ]

        self.second_day_rows = [
            ["Date:", self.second_day_datestring],
            [
                "Meal",
                "Brand",
                "Name",
                "Servings",
                "Calories",
                "Nutrient A",
                "Nutrient B",
                "Nutrient C",
            ],
            ["breakfast", None, "cereal", 1, 220, 7, "2mg", 120],
            ["breakfast", "A2", "milk", 1, 140, 22, 0.3, 130],
            ["lunch", None, "veggie pizza (slice)", 2, 530, 15, 29, 34],
            ["snacks", "Foo", "protein bar", 1, 220, 25, 5, 44],
            ["dinner", None, "salad", 1, 120, 3, 32, 22],
            ["dinner", None, "salmon", 1, 380, 36, 90, 88],
        ]

        wb = xlwt.Workbook()
        ws = wb.add_sheet("Test Sheet")

        write_sheet_data(
            ws, [*self.first_day_rows, *filler_rows, *self.second_day_rows]
        )

        wb.save(XLS_TEST_FILENAME)

    def test_valid_dataframe_from_file(self):
        meals_df = mpp().get_meals_df(XLS_TEST_FILENAME)
        self.assertIsInstance(meals_df, pd.DataFrame)

    def test_date_values(self):
        meals_df = mpp().get_meals_df(XLS_TEST_FILENAME)
        expected_date_values = [self.first_day_datestring, self.second_day_datestring]
        self.assertCountEqual(meals_df["Date"].unique().tolist(), expected_date_values)

    def test_rows(self):
        meals_df = mpp().get_meals_df(XLS_TEST_FILENAME)
        expected_rows = [
            [self.first_day_datestring] + row for row in self.first_day_rows[2:]
        ] + [[self.second_day_datestring] + row for row in self.second_day_rows[2:]]
        actual_rows = (
            meals_df.loc[
                meals_df["Date"].isin(
                    [self.first_day_datestring, self.second_day_datestring]
                )
            ]
            .replace({np.nan: None})
            .values.tolist()
        )
        self.assertCountEqual(expected_rows, actual_rows)

    def tearDown(self):
        testfile_cleanup_helper(XLS_TEST_FILENAME)


if __name__ == "__main__":
    unittest.main()
