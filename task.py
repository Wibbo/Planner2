import pandas as pd
import datetime as dt


class DAITask:

    @staticmethod
    def get_age_from_date(row, start_default=5):
        """
        Calculates the number of days that a non-complete task has existed.
        :param row: A row from the task spreadsheet.
        :param start_default: A default value for the start date if none exists.
        :return: The age of a task in days.
        """
        start_date = row['Created Date']

        if row.Progress == 'Completed':
            end_date = row['Completed Date']
            start_date = row['Start Date']
        else:
            end_date = dt.date.today()
            start_date = row['Created Date']

        if pd.isnull(start_date):
            start_date = end_date - dt.timedelta(start_default)

        diff = end_date - start_date
        return diff.days

    @staticmethod
    def get_owner(description):
        """
        Determines if an owner is specified in the task description, returning it if it exists.
        :param description: The contents of the task's description field.
        :return: The name of the owner if one is found.
        """
        data = description.split()
        index = 0
        first = ''
        last = ''

        try:
            for word in data:
                if word == 'Owner:':
                    first = data[index + 1]
                    last = data[index + 2]
                    break
                index += 1
        except:
            return 'Unspecified'
        finally:
            return first + ' ' + last

    @staticmethod
    def reformat_date_field(date_field):
        """
        Changes the format of date fields so that they can be processed.
        :return: The newly formatted date.
        """
        date_field = pd.to_datetime(date_field, format='%d/%m/%Y')
        date_field = date_field.dt.date
        return date_field

    def split_tasks(self):
        """
        Splits the Completed Checklist Items column into steps done and total steps fields.
        :return: Nothing
        """
        steps = self.df_tasks['Completed Checklist Items'].str.split("/", expand=True)
        self.df_tasks[['Steps Done', 'Steps Total']] = steps
        self.df_tasks['Steps Done'] = pd.to_numeric(self.df_tasks['Steps Done'], downcast='integer')
        self.df_tasks['Steps Total'] = pd.to_numeric(self.df_tasks['Steps Total'], downcast='integer')

    def __init__(self, excel_file):
        """
        Constructor for the DAITask class.
        :param excel_file: The name of the Excel file extract from Planner to process.
        :type excel_file:
        """

        # Read the task spreadsheet into a pandas dataframe.
        self.df_tasks = pd.read_excel(excel_file, sheet_name="Tasks", skiprows=4, usecols="A:Q")

        self.df_tasks['Created Date'] = self.reformat_date_field(self.df_tasks['Created Date'])
        self.df_tasks['Start Date'] = self.reformat_date_field(self.df_tasks['Start Date'])
        self.df_tasks['Due Date'] = self.reformat_date_field(self.df_tasks['Due Date'])
        self.df_tasks['Completed Date'] = self.reformat_date_field(self.df_tasks['Completed Date'])

        self.split_tasks()

        self.df_tasks['Description'] = self.df_tasks['Description'].astype(str)

        self.df_tasks['Owner'] = self.df_tasks['Description'].apply(self.get_owner)
        # self.df_tasks['Status'] = self.df_tasks['Description'].apply(self.get_status())
        self.df_tasks['Age'] = self.df_tasks.apply(self.get_age_from_date, axis=1)

        self.df_tasks = self.df_tasks.drop(['Task ID', 'Completed By', 'Description',
                                           'Completed Checklist Items', 'Checklist Items'],
                                           axis=1)


