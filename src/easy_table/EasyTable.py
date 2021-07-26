# EasyTable class
class EasyTable(object):
    # Constructor
    def __init__(self, table_name=""):
        self.table_name = table_name
        self.table_structure = {
            "top-left-corner": "",
            "top-right-corner": "",
            "bottom-left-corner": "",
            "bottom-right-corner": "",
            "vertical-outer": "",
            "horizontal-outer": "",
            "vertical-inner": "",
            "horizontal-inner": "",
            "intersection-inner": ""
        }
        self.data = []
        self.table_columns = []
        self.table_column_lengths = {}
        self.table_data = []

    # Represent the object
    def __repr__(self):
        return "<EasyTable {}>".format(self.table_name)

    # Set the table name
    def setTableName(self, table_name):
        self.table_name = table_name

    # Set the corners of the table
    def setCorners(self, top_left, top_right, bottom_left, bottom_right):
        self.table_structure["top-left-corner"] = top_left
        self.table_structure["top-right-corner"] = top_right
        self.table_structure["bottom-left-corner"] = bottom_left
        self.table_structure["bottom-right-corner"] = bottom_right

    # Set the outer structure of the table
    def setOuterStructure(self, vertical_outer, horizontal_outer):
        self.table_structure["vertical-outer"] = vertical_outer
        self.table_structure["horizontal-outer"] = horizontal_outer

    # Set the innser structure of the table
    def setInnerStructure(self, vertical_inner, horizontal_inner,
                          intersection_inner):
        self.table_structure["vertical-inner"] = vertical_inner
        self.table_structure["horizontal-inner"] = horizontal_inner
        self.table_structure["intersection-inner"] = intersection_inner

    # Display the table structure
    def displayTableStructure(self):
        print("----TABLE PROPERTIES----")
        print("Table Corners: " + self.table_structure["top-left-corner"] +
              " " + self.table_structure["top-right-corner"] + " " +
              self.table_structure["bottom-left-corner"] + " " +
              self.table_structure["bottom-right-corner"])
        print("Table Outer: " + self.table_structure["vertical-outer"] + " " +
              self.table_structure["horizontal-outer"])
        print("Table Inner: " + self.table_structure["vertical-inner"] + " " +
              self.table_structure["horizontal-inner"] + " " +
              self.table_structure["intersection-inner"])

# #############################################################################################
# TABLE SETTING FUNCTIONS
# #############################################################################################

    # Set the data for the table
    # Data can be a list of lists or a lists of dicts
    # If using a list of lists, setColumns needs to be called before
    def setData(self, data):
        self.data = data
        self.rollbackData()
        self.findColumns()

# #############################################################################################
# TABLE FORMATTING FUNCTIONS
# #############################################################################################

    # findColumns is used for getting column names when using a list of dicts
    # Iterate through items in the list to identify the keys that will be
    # columns
    def findColumns(self):
        for row in self.table_data:
            if type(row) is dict:
                for column in row.keys():
                    if column not in self.table_columns:
                        self.table_columns.append(str(column))

    # setColumns is used to set the columns, it takes a list
    # Set columns manually
    def setColumns(self, columns):
        self.table_columns = []
        for column in columns:
            if column not in self.table_columns:
                self.table_columns.append(str(column))

    # Takes a dict of column names to add
    # Just adds it to each row, with empty data
    def addColumns(self, columns):
        for key, value in columns.items():
            for row in self.table_data:
                row[key] = value

        if(len(self.table_columns) != 0):
            self.findColumns()

    # Takes a list of column names to add
    # Just removes the columns from each row
    def removeColumns(self, columns):
        for column in columns:
            for row in self.table_data:
                try:
                    del row[column]
                except KeyError:
                    pass

            try:
                self.table_columns.remove(column)
            except ValueError:
                pass

        if(len(self.table_columns) != 0):
            self.findColumns()

    # Format the data so each row is a dictionary with the data
    def formatTableData(self):
        temp_data = []
        # If there is no data dont do anything
        if len(self.table_data) == 0:
            pass
        # If the type is a dictionary
        if type(self.table_data[0]) is dict:
            # Go through each row
            for row in self.table_data:
                formatted_row = {}
                # Go through each key
                for key in self.table_columns:
                    # If it is not there add in an empty string
                    try:
                        formatted_row[key] = row[key]
                    except KeyError:
                        formatted_row[key] = ""
                temp_data.append(formatted_row)
        # If the type is a list
        elif type(self.data[0]) is list:
            # Go through each row
            for row in self.table_data:
                formatted_row = {}
                # Go through each item in the row
                for x in range(0, len(self.table_columns)):
                    # If it is not there add an empty string
                    try:
                        formatted_row[self.table_columns[x]] = row[x]
                    except IndexError:
                        formatted_row[self.table_columns[x]] = ""
                temp_data.append(formatted_row)
        self.table_data = temp_data

# ###############################################################################################
# TABLE DISPLAY FUCNTIONS
# ###############################################################################################

    # Get the max length in each item
    def getMaxLengthPerColumn(self):
        self.table_column_lengths = {}
        # Loop throught each column
        for column in self.table_columns:
            self.table_column_lengths[column] = len(str(column).encode('gbk'))
            # Go through each row and find the max length
            for row in self.table_data:
                if len(str(row[column]).encode('gbk')) > self.table_column_lengths[column]:
                    self.table_column_lengths[column] = len(str(row[column]).encode('gbk'))

    # Print the table header
    def printTableHeader(self):
        lengths = 0
        for length in self.table_column_lengths.values():
            lengths += length

        # Print the top row of the header
        row_string = self.table_structure["top-left-corner"]
        row_string += self.table_structure["horizontal-outer"] * (
            lengths + len(self.table_column_lengths) - 1)
        row_string += self.table_structure["top-right-corner"]
        print(row_string)

        row_string = self.table_structure["vertical-outer"]
        # Loop though and print each column name
        for column_num, column_name in enumerate(self.table_columns):
            column_string = ""
            column_string += " " * int(
                ((self.table_column_lengths[column_name] -
                  len(str(self.table_columns[column_num]).encode('gbk'))) / 2))
            column_string += self.table_columns[column_num]
            column_string += " " * (
                self.table_column_lengths[column_name] - len(column_string.encode('gbk')))
            row_string += column_string
            if(column_num != len(self.table_columns) - 1):
                row_string += self.table_structure["vertical-inner"]
        row_string += self.table_structure["vertical-outer"]
        print(row_string)

        self.printDividerRow()

    # Print the data for each row
    def printDataRow(self, row):
        row_string = self.table_structure["vertical-outer"]
        # Loop through each column, get the length
        # and use it to print the formatted column data
        for column_num, column_name in enumerate(self.table_columns):
            column_string = ""
            column_string += " " * int(
                ((self.table_column_lengths[column_name] -
                  len(str(row[self.table_columns[column_num]]).encode('gbk'))) / 2))
            column_string += str(row[self.table_columns[column_num]])
            column_string += " " * (
                self.table_column_lengths[column_name] - len(column_string.encode('gbk')))
            row_string += column_string
            if(column_num != len(self.table_columns) - 1):
                row_string += self.table_structure["vertical-inner"]
        row_string += self.table_structure["vertical-outer"]
        print(row_string)

    # Print the divider for the row
    def printDividerRow(self):
        row_string = self.table_structure["vertical-outer"]
        # Loop through each column to get the
        # length of each column and print the divider
        for column_num, column_name in enumerate(self.table_columns):
            row_string += self.table_structure[
                "horizontal-inner"] * self.table_column_lengths[column_name]
            if(column_num != len(self.table_columns) - 1):
                row_string += self.table_structure["intersection-inner"]
        row_string += self.table_structure["vertical-outer"]
        print(row_string)

    # Print the foot of the table
    def printTableFooter(self):
        lengths = 0
        # Loop through the lengths of the columns
        for length in self.table_column_lengths.values():
            lengths += length

        row_string = self.table_structure["bottom-left-corner"]
        row_string += self.table_structure["horizontal-outer"] * (
            lengths + len(self.table_column_lengths) - 1)
        row_string += self.table_structure["bottom-right-corner"]
        print(row_string)

    # Display the table
    def displayTable(self):
        # Get the max length
        self.getMaxLengthPerColumn()
        # Print the header
        self.printTableHeader()
        # Print the rows
        for row_num, row in enumerate(self.table_data):
            self.printDataRow(row)
            if(row_num != len(self.table_data) - 1):
                self.printDividerRow()
        self.printTableFooter()

# ##############################################################################################
# DATA ALTERING FUNCTIONS
# ##############################################################################################

    # Sorta data based off the given condition
    def sortData(self, condition=None, reverse=False):
        self.table_data.sort(key=condition, reverse=reverse)

    # Store table data in background "checkpoint"
    def commitData(self):
        self.data = self.table_data.copy()

    # Roll back the table data to the previous "checkpoint"
    def rollbackData(self):
        self.table_data = self.data.copy()

        # If the table columns has already been set
        # then format the data
        if(len(self.table_columns) != 0):
            self.formatTableData()

    # Insert data into the table
    def insertData(self, new_data):
        # If the new data it a list
        if type(new_data) is list:
            # Loop through the new data
            for row in new_data:
                # If the new row is a dict then add it to the data
                if type(row) is dict:
                    self.table_data.append(row)
                if type(row) is list:
                    formatted_row = {}
                    for column_num, column_name in enumerate(
                            self.table_columns):
                        try:
                            formatted_row[column_name] = row[column_num]
                        except IndexError:
                            formatted_row[column_name] = ""
                    self.table_data.append(formatted_row)
        # If the new data is a dict then add it to the data
        elif type(new_data) is dict:
            self.table_data.append(new_data)

        # If there are columns find the new columns
        if(len(self.table_columns) != 0):
            self.findColumns()
            self.formatTableData()

    # Delete data
    # Pass a function that returns true if you dont want it removed
    def deleteData(self, condition=None):
        remaining = []

        # If there is a condition
        if condition is not None:
            # Go through each row and if the
            # condition is true add it to a
            # remaining list
            for row in self.table_data:
                if condition(row):
                    remaining.append(row)

        self.table_data = remaining

    # Update data
    # Pass a function that returns the new value
    # Pass a function that returns true if it you want it updated
    def updateData(self, value, condition=None):
        # Loop through the table data
        # If there is a condition function
        # use it, use value to set the row
        for row in self.table_data:
            if condition is not None:
                if condition(row):
                    row = value(row)
            else:
                row = value(row)
