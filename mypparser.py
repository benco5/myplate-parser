"""
Livestrong MyPlate Detailed Export Parser module.
"""

import warnings
import pandas as pd
import olefile

warnings.filterwarnings(action="ignore", category=FutureWarning)


class MyPlateDetailedExportParser:

    """
    Parses Livestrong MyPlate app detailed export files and extracts meal data.

    This class provides methods to parse and extract meal data from MyPlate
    detailed-level export files in .xls format. It addresses the issue of
    'Workbook corruption' encountered when using the xlrd engine to read the
    files. The class includes a method to convert the .xls file to a raw DataFrame
    and another method to extract and transform the meal data into a final DataFrame.
    The extracted meal data includes information such as dates, meals, and associated
    nutritional values. The class allows users to retrieve the final DataFrame
    containing the extracted meal data from a given .xls file.

    Note: The class assumes the .xls files follow the specific structure and format
    used by MyPlate for detailed-level exports.
    """

    MEAL_NAMES = ["breakfast", "lunch", "dinner", "snacks"]

    def __init__(self):
        self.path = None
        self.raw_df = None
        self.meals_final_df = None

    def xls_to_raw_df(self):
        """
        Converts the MyPlate detailed-level export .xls file to a raw DataFrame.

        This method is necessary because the xlrd engine used by pandas' read_excel
        method (or if xlrd is used directly) raises an exception when attempting to
        read MyPlate detailed-level export .xls files:

        xlrd.compdoc.CompDocError: Workbook corruption: seen[2] == 4

        The following solution provides a workaround for this issue, inspired by:
        https://stackoverflow.com/a/60416081/3761560

        Prior attempts to read the MyPlate .xls file with xlrd directly using the
        'ignore_workbook_corruption=True' option only enabled reading up to row 163
        for multiple files. However, the actual data appeared intact, suggesting that
        the reported "corruption" is benign and likely an artifact of how the .xls
        file was generated or a limitation of xlrd, depending on the perspective.
        This issue seems to be related to "Compound File Binary" compatibility.

        For further reference, see:
        https://web.archive.org/web/20190311101348/http://www.crimulus.com/2013/09/19/reading-compound-file-binary-format-files-generated-by-phpexcel-with-pythondjango-xlrd/
        """
        with open(self.path, "rb") as file:
            ole = olefile.OleFileIO(file)
            if ole.exists("Workbook"):
                d = ole.openstream("Workbook")
                self.raw_df = pd.read_excel(d, engine="xlrd", header=None)

    def extract_and_transform_meals(self):
        """
        Extracts and transforms meals data from the raw DataFrame.

        The function iterates over the rows of the raw DataFrame and extracts meal data
        based on specific column labels. It creates a new DataFrame with the extracted
        meal data and sets it as the 'meals_final_df' attribute of the object.
        """
        df_raw = self.raw_df
        df_meals = pd.DataFrame()
        curr_date = None
        curr_meal_ref_row = None

        for i in range(len(df_raw)):
            # First column value as "label" to determine nature of column data
            label = df_raw.iloc[i, 0]

            # Set current date
            if label == "Date:":
                curr_date = df_raw.iloc[i, 1]

            # Extract / create meal reference row, i.e., column heading values
            if label == "Meal":
                curr_meal_ref_row = df_raw.iloc[i]
                curr_meal_ref_row = ["Date"] + list(curr_meal_ref_row)

            # Extract / create meal item rows
            if label in self.MEAL_NAMES:
                curr_meal_value_row = [curr_date] + list(df_raw.iloc[i])
                new_row_dict = dict(zip(curr_meal_ref_row, curr_meal_value_row))
                df_meals = df_meals.append(new_row_dict, ignore_index=True)

        self.meals_final_df = df_meals

    def get_meals_df(self, xls_path):
        """
        Returns a DataFrame containing extracted meal data from the provided .xls file.

        :param xls_path: The file path of the .xls file.
        :return: DataFrame containing meal data.
        """
        self.path = xls_path
        self.xls_to_raw_df()
        self.extract_and_transform_meals()
        return self.meals_final_df
