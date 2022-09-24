
import csv
from enum import Enum
from datetime import datetime
from tkinter import E


class CATEGORY(Enum):
    HOUSING = "Housing"
    ENTERTAINMENT = "Entertainment"
    SHOPPING = "Shopping"
    FOOD = "Food"
    INSURANCE = "Insurance"
    UTILITIES = "Utilities"
    TRANSPORTATION = "Transportation"
    PERSONAL = "Personal"
    DEBT = "Debt"
    INCOME = "Income"
    OTHER = "Other"
    SUBSCRIPTION = "Subscription"
    HOLIDAYS = "Holidays"
    TAX = "Tax"


class Expense():

    def __init__(self, date, item, debit, credit):
        self.date = date
        self.item = item
        self.debit = debit
        self.credit = credit
        self.category = self.determine_category()

    def determine_category(self):

        item = self.item.lower().strip()

        if item in ["water", "electricity","phone", "internet"]:
            return CATEGORY.UTILITIES
        elif item in ["train", "subway"] or "car" in item:
            return CATEGORY.TRANSPORTATION
        elif item in ["bar", "entertainment"]:
            return CATEGORY.ENTERTAINMENT
        elif item in ["groceries", "restaurant"]:
            return CATEGORY.FOOD
        elif item in ["salary","income","transfer"]:
            return CATEGORY.INCOME
        elif "subscription" in item:
            return CATEGORY.SUBSCRIPTION
        elif item == "rent":
            return CATEGORY.HOUSING
        elif item in ["tax", "contribution"]:
            return CATEGORY.TAX
        elif item == "insurance":
            return CATEGORY.INSURANCE
        elif item == "holidays":
            return CATEGORY.HOLIDAYS
        elif item == "shopping":
            return CATEGORY.SHOPPING
        elif item in ["repayment", "debt"]:
            return CATEGORY.DEBT
        elif item == "other":
            return CATEGORY.OTHER


def get_data(file):
    """
    Extract the financial data from a csv file 

    Arguments:
    - file: csv file which contains all the data 

    Returns:
    - data
    """

    expenses = []

    # read csv file
    with open(file) as file:
        
        # csv reader
        csv_reader = csv.reader(file, delimiter=";")

        # get headers and rows
        headers = next(csv_reader)
        rows = [[item.strip() for item in row] for row in csv_reader]
       
        # sort rows by date
        rows.sort(key=lambda row: datetime.strptime(row[0], "%d/%m/%Y"))

        # add headers in list
        headers.insert(2,"Category")
        expenses.append(headers)

        # add rows in list
        for row in rows:
            
            # Expense object
            expense_obj = Expense(row[0], row[1], row[2], row[3])
            
            # add data in list
            try:
                expenses.append([expense_obj.date,
                            expense_obj.item,
                            expense_obj.category.value,
                            expense_obj.debit,
                            expense_obj.credit])
            except Exception as e:
                raise e
                

    return expenses

