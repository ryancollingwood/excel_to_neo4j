# -*- coding: utf-8 -*-
"""
Created on Mon Jan 9 10:59:37 2017

Script for importing tabular excel document into a Neo4j graph database.
The intension is to import not only cell value but the additional information, such as:
* Formulas
* Cell Decorations - Colour, Borders
* Font Decorations - Colour, Family, Modifiers

Neo4j - https://neo4j.com
Neo4j Python Driver Soruce - https://github.com/neo4j/neo4j-python-driver
xLwings - https://www.xlwings.org/
xlwings Source - https://github.com/ZoomerAnalytics/xlwings

@author: ryancollingwood@gmail.com

export_sheet is the entry point function

Some quick improvements that could be made:

* Move the consts to an external configuration file;
* Find a better solution to reading through blank rows, rather using stopping after C_MAX_EMPTY blank rows;
* Use Pandas Dataframe for creating an in memory representation of the Spreadsheet;
* Use Transaction for handling upserts, so that rollback is possible;

"""

import xlwings as xw
import datetime
from neo4j.v1 import GraphDatabase, basic_auth

C_VALUE_TEMP = "{ value: {value} }"
C_MAX_EMPTY = 100
C_USER_LOGIN = "neo4j"
C_USER_PASSWORD = "neo4j"
C_SERVER_URL = "bolt://localhost:7687"
C_DATA_START_ROW = 2

def load_work_book(fileName):
    book = xw.Book(fileName)
    return book

def get_work_book_sheet(workBook, sheetIndex):
    sheet = workBook.sheets[sheetIndex]
    return sheet

# naive function to make a string into one cypher will be happy with
def to_neo_label(value):
    # TODO: Regex Find and Replace
    result = str(value).upper().strip().replace(" ", "_").replace("-", "_")
    return result

# make a string into NeoPropertyName
def to_neo_property_name(value):
    result = to_neo_label(value)
    result = result.replace("_", " ").title()
    result = result.replace(" ", "")
    return result

# very ugly isNumeric check - excel stores all number as flaots
def is_numeric(value):
    try:
        float(value)
        return True
    except ValueError:
        # wasn't numeric
        pass
        return False

    return False

 # naive is datetime test
def is_datetime(value):
    try:
        if (isinstance(value, datetime.datetime)):
            return True
        else:
            return False
    except ValueError:
        # not datetime
        pass
        return False

    return False

# return a set of columnHeaders
# limitations will not cope well with merged cells
# may need to perhaps return a dict that stores the column index
def extract_column_headers(rowData, columnCount):
    result = []

    noneColCount = 0

    for column in range(1, columnCount):
        #have we had more blanks than we wanted?
        if (noneColCount > C_MAX_EMPTY):
            break

        # now we have a cell
        value = rowData[column].value

        if value:
            print(value)
        else:
            noneColCount = noneColCount + 1
            continue

        result.append(to_neo_label(value))

    return result

# return a dict of row values
def read_row(rowData, columnHeaders):
    result = {}

    # for each column - starting at 1 as excel is 1 based index
    noneColCount = 0

    for column in range(1, len(columnHeaders)):
        #have we had more blanks than we wanted?
        if (noneColCount > C_MAX_EMPTY):
            break

        # now we have a cell
        value = rowData[column].value

        if not value:
            noneColCount = noneColCount + 1
            continue

        # TODO: transformations ala magic regex?

        # column - 1 as python array are 0 indexed
        result[columnHeaders[column-1]] = value

    return result

# merge a node
def neo_merge_node(label, importLabel, value, session):
    cypher = neo_node_cypher("n", label, importLabel)
    query = " ".join( ["MERGE", cypher] )
    # session.run(query, {"value": value})

    with session.begin_transaction() as tx:
        tx.run(query, {"value": value})
        tx.success = True


# get a cypher string for matching/creating/merging a node
def neo_node_cypher(selector, label, importLabel, value_temp = C_VALUE_TEMP):
    template = "({selector} :{label} :{importLabel} {value_temp})"
    template = template.format(selector = selector, label = label, importLabel = importLabel, value_temp = value_temp)

    return template

# get a cypher string for matching/creating/merging a relationship
def neo_relationship_cypher(selector, label, value_temp = C_VALUE_TEMP):
    template = "[{selector} :`{label}` {value_temp}]"
    template = template.format(selector = selector, label = label, value_temp = value_temp)
    return template

# take a python dict and convert it into a string for use in a parameterised cypher query
def dict_to_cypher_params(value):
    result = ""
    for key in value:
        if (result != ""):
            result = "".join([result, ", "])

        result = "".join([result, key, ": ", "{", key, "}"])

    if (result != ""):
        result = "".join(["{", result, "}"])

    return result

# create relationships between cells in a row, which are now nodes
def neo_create_relationships(rowData, forKey, relationshipProperties, session):
    fromNode = neo_node_cypher("a", forKey, "test", "{ value: {a_value} }")

    for key in rowData:
        if (key == forKey):
            continue

        toNode = neo_node_cypher("b", key, "test", "{ value: {b_value} }")

        # get the relationship name as left to right from (a) to (b)
        relName = "_".join([forKey, key])
        relvalues = dict_to_cypher_params(relationshipProperties)
        relCypher = neo_relationship_cypher("r", relName, relvalues)

        query = "MATCH {a}, {b} WITH a,b CREATE (a)-{relCypher}->(b)"
        query = query.format(a = fromNode, b = toNode, relCypher = relCypher)

        # update query params to include values to match from (a) to (b) node
        queryParams = relationshipProperties.copy()
        queryParams.update({"a_value": rowData[forKey], "b_value": rowData[key]})

        # session.run(query, queryParams)
        with session.begin_transaction() as tx:
            tx.run(query, queryParams)
            tx.success = True


# foreach cell in a row in the spreadsheet create nodes for the categorical data
# then link the catagorical nodes with relationsiphs containing the numeric datatypes
# found in the row
def export_rows(columnHeaders, startDataRow, sheet, driver):

    # keep track of the number of empty rows
    # otherwise this will run till the max rowsize in an excel document
    noneRowCount = 0

    # for each row - starting at 1 as excel is 1 based index
    for row in range(startDataRow, sheet.cells.rows.count):

        #do we want to keep scanning rows?
        #chceck firest cell of this row
        cell = sheet.range((row, 1))
        if (cell.value == None):
            noneRowCount = noneRowCount + 1

        if (noneRowCount > C_MAX_EMPTY):
            break

        rowData = read_row(sheet.cells.rows(row), columnHeaders)

        if (len(rowData) == 0):
            continue

        # create a dictionary for storing relationship properties
        relationshipProperties = {}

        # get a session
        session = driver.session()
        try:
            for key in rowData:
                # numeric datatypes we want inside relationships
                if (is_numeric(rowData[key])):
                    relationshipProperties[key] = rowData[key]
                else:
                    # neo4j doesn't like python datetime storing as ISO datetime
                    # TODO: also store as epoch as that is neo4j's datetime
                    # TODO: maybe build nodes for Year, Month, Day and link em
                    if (is_datetime(rowData[key])):
                        rowData[key] = rowData[key].isoformat()

                    neo_merge_node(key, "test", rowData[key], session)

            # remove the values in a our relationships from rowData
            for key in relationshipProperties:
                del rowData[key]

            # now relationships our nodes
            for key in rowData:
                neo_create_relationships(rowData, key, relationshipProperties, session)
        finally:
            # close session and return to thread pool
            session.close()
			
	return

# from excel to neo4j
# this is the entry point function
def export_sheet(fileName, sheetIndex):
    book = load_work_book(fileName)
    sheet = get_work_book_sheet(book, sheetIndex)

    columnHeaders = extract_column_headers(sheet.cells.rows(1), sheet.cells.columns.count)

    startDataRow = C_DATA_START_ROW

    driver = GraphDatabase.driver(C_SERVER_URL, auth = basic_auth(C_USER_LOGIN, C_USER_PASSWORD))

    export_rows(columnHeaders, startDataRow, sheet, driver)
	
	return
