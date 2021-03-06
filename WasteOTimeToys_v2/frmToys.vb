﻿'------------------------------------------------------------
'-                  File Name: frmToys.vb                   -
'-                 Part of Project: Assign9                 -
'------------------------------------------------------------
'-                Written By: Elijah Wilson                 -
'-                  Written On: 04/05/2016                  -
'------------------------------------------------------------
'- File Purpose:                                            -
'-                                                          -
'- This is the main form for the application. It handles    -
'- all of the form funtionality for the program.            -
'------------------------------------------------------------
'- Program Purpose:                                         -
'-                                                          -
'- To generate an Excel spreadsheet based on the input data -
'- file with many statistics.                               -
'------------------------------------------------------------
'- Global Variable Dictionary (alphabetically):             -
'- (None)                                                   -
'------------------------------------------------------------
Imports Microsoft.Office.Interop

Public Class frmToys
    Private Const FILE_NAME As String = "ToyOrder.txt"

    ' Column Letters
    Private Const COL_FIRST_NAME As String = "A"
    Private Const COL_LAST_NAME As String = "B"
    Private Const COL_ORDER_ID As String = "C"
    Private Const COL_EMP_ID As String = "D"
    Private Const COL_GAME_SALES As String = "F"
    Private Const COL_DOLL_SALES As String = "G"
    Private Const COL_BUILD_SALES As String = "H"
    Private Const COL_MODEL_SALES As String = "I"
    Private Const COL_TOTAL_SALES As String = "J"
    Private Const COL_MIN_SALES As String = "K"
    Private Const COL_AVG_SALES As String = "L"
    Private Const COL_MAX_SALES As String = "M"
    Private Const COL_GAMES_QTY As String = "O"
    Private Const COL_DOLLS_QTY As String = "P"
    Private Const COL_BUILD_QTY As String = "Q"
    Private Const COL_MODEL_QTY As String = "R"
    Private Const COL_TOTAL_QTY As String = "S"
    Private Const COL_MIN_QTY As String = "T"
    Private Const COL_AVG_QTY As String = "U"
    Private Const COL_MAX_QTY As String = "V"
    Private Const COL_AGG_TITLES As String = "E"

    ' As Integers
    Private INT_COL_FIRST_NAME As Integer = ColumnLetterToColumnIndex(COL_FIRST_NAME)
    Private INT_COL_LAST_NAME As Integer = ColumnLetterToColumnIndex(COL_LAST_NAME)
    Private INT_COL_ORDER_ID As Integer = ColumnLetterToColumnIndex(COL_ORDER_ID)
    Private INT_COL_EMP_ID As Integer = ColumnLetterToColumnIndex(COL_EMP_ID)
    Private INT_COL_GAME_SALES As Integer = ColumnLetterToColumnIndex(COL_GAME_SALES)
    Private INT_COL_DOLL_SALES As Integer = ColumnLetterToColumnIndex(COL_DOLL_SALES)
    Private INT_COL_BUILD_SALES As Integer = ColumnLetterToColumnIndex(COL_BUILD_SALES)
    Private INT_COL_MODEL_SALES As Integer = ColumnLetterToColumnIndex(COL_MODEL_SALES)
    Private INT_COL_TOTAL_SALES As Integer = ColumnLetterToColumnIndex(COL_TOTAL_SALES)
    Private INT_COL_MIN_SALES As Integer = ColumnLetterToColumnIndex(COL_MIN_SALES)
    Private INT_COL_AVG_SALES As Integer = ColumnLetterToColumnIndex(COL_AVG_SALES)
    Private INT_COL_MAX_SALES As Integer = ColumnLetterToColumnIndex(COL_MAX_SALES)
    Private INT_COL_GAMES_QTY As Integer = ColumnLetterToColumnIndex(COL_GAMES_QTY)
    Private INT_COL_DOLLS_QTY As Integer = ColumnLetterToColumnIndex(COL_DOLLS_QTY)
    Private INT_COL_BUILD_QTY As Integer = ColumnLetterToColumnIndex(COL_BUILD_QTY)
    Private INT_COL_MODEL_QTY As Integer = ColumnLetterToColumnIndex(COL_MODEL_QTY)
    Private INT_COL_TOTAL_QTY As Integer = ColumnLetterToColumnIndex(COL_TOTAL_QTY)
    Private INT_COL_MIN_QTY As Integer = ColumnLetterToColumnIndex(COL_MIN_QTY)
    Private INT_COL_AVG_QTY As Integer = ColumnLetterToColumnIndex(COL_AVG_QTY)
    Private INT_COL_MAX_QTY As Integer = ColumnLetterToColumnIndex(COL_MAX_QTY)
    Private INT_COL_AGG_TITLES As Integer = ColumnLetterToColumnIndex(COL_AGG_TITLES)

    Private Const STARTING_ROW As Integer = 1

    ' Employee Formulas - String.Format with row
    Private Const FORMULA_TOTAL_SALES As String = "=sum(" & COL_GAME_SALES & "{0}.." & COL_MODEL_SALES & "{0})"
    Private Const FORMULA_MIN_SALES As String = "=min(" & COL_GAME_SALES & "{0}.." & COL_MODEL_SALES & "{0})"
    Private Const FORMULA_AVG_SALES As String = "=average(" & COL_GAME_SALES & "{0}.." & COL_MODEL_SALES & "{0})"
    Private Const FORMULA_MAX_SALES As String = "=max(" & COL_GAME_SALES & "{0}.." & COL_MODEL_SALES & "{0})"
    Private Const FORMULA_TOTAL_QTY As String = "=sum(" & COL_GAMES_QTY & "{0}.." & COL_MODEL_QTY & "{0})"
    Private Const FORMULA_MIN_QTY As String = "=min(" & COL_GAMES_QTY & "{0}.." & COL_MODEL_QTY & "{0})"
    Private Const FORMULA_AVG_QTY As String = "=average(" & COL_GAMES_QTY & "{0}.." & COL_MODEL_QTY & "{0})"
    Private Const FORMULA_MAX_QTY As String = "=max(" & COL_GAMES_QTY & "{0}.." & COL_MODEL_QTY & "{0})"
    Private Const FORMULA_AGGREGATE As String = "({0}{1}..{0}{2})"

    Private maxRowWithData As Integer = STARTING_ROW + 1  ' in case no data

    '------------------------------------------------------------
    '-             Subprogram Name: btnMagic_Click              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 04/05/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Handles when the main button is clicked. Parses the      -
    '- input file and generates the excel report.               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - The object that raised the event                -
    '- e - The EventArgs sent with the event                    -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- checkExcel - Possibly an Excel.Application if one        -
    '-              already exists                              -
    '- employees - A List of Employees that were parsed from    -
    '-             the input file                               -
    '- excelDoc - The Excel.Application object that is the main -
    '-            driver for the application                    -
    '- parser - A Parser object that parses the input file      -
    '------------------------------------------------------------
    Private Sub btnMagic_Click(sender As Object, e As EventArgs) Handles btnMagic.Click
        Dim checkExcel As Object
        Dim excelDoc As Excel.Application
        Dim parser As New Parser(FILE_NAME)
        Dim employees As List(Of Employee) = (From emp In parser.parseFile()
                                              Order By emp.id Ascending
                                              Select emp).ToList()

        Try
            checkExcel = GetObject(, "Excel.Application")
        Catch ex As Exception
            ' Excel not running
        End Try

        If checkExcel Is Nothing Then
            excelDoc = New Excel.Application()
        Else
            excelDoc = checkExcel
        End If

        MessageBox.Show("Generating spreadsheet...")

        excelDoc.Workbooks.Add()
        excelDoc.Sheets.Add()

        writeHeaders(excelDoc)
        writeEmployees(excelDoc, employees)
        writeAggregateRows(excelDoc)
        excelDoc.Visible = True

        MessageBox.Show("Pausing...")
        excelDoc.Quit()
        excelDoc = Nothing
    End Sub

    '------------------------------------------------------------
    '-              Subprogram Name: writeHeaders               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 04/05/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Writes the Column headers to the excel spreadsheet       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- excelDoc - The Excel.Application to be used within the   -
    '-            Subroutine                                    -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub writeHeaders(ByRef excelDoc As Excel.Application)
        excelDoc.Cells(STARTING_ROW, INT_COL_FIRST_NAME) = "First Name"
        excelDoc.Cells(STARTING_ROW, INT_COL_LAST_NAME) = "Last Name"
        excelDoc.Cells(STARTING_ROW, INT_COL_ORDER_ID) = "Order ID"
        excelDoc.Cells(STARTING_ROW, INT_COL_EMP_ID) = "Employee ID"

        excelDoc.Cells(STARTING_ROW, INT_COL_GAME_SALES) = "Games Sales"
        excelDoc.Cells(STARTING_ROW, INT_COL_DOLL_SALES) = "Dolls Sales"
        excelDoc.Cells(STARTING_ROW, INT_COL_BUILD_SALES) = "Build Sales"
        excelDoc.Cells(STARTING_ROW, INT_COL_MODEL_SALES) = "Model Sales"
        excelDoc.Cells(STARTING_ROW, INT_COL_TOTAL_SALES) = "Total Sales"
        excelDoc.Cells(STARTING_ROW, INT_COL_MIN_SALES) = "Min Sales"
        excelDoc.Cells(STARTING_ROW, INT_COL_AVG_SALES) = "Avg Sales"
        excelDoc.Cells(STARTING_ROW, INT_COL_MAX_SALES) = "Max Sales"

        excelDoc.Cells(STARTING_ROW, INT_COL_GAMES_QTY) = "Games Qty"
        excelDoc.Cells(STARTING_ROW, INT_COL_DOLLS_QTY) = "Dolls Qty"
        excelDoc.Cells(STARTING_ROW, INT_COL_BUILD_QTY) = "Build Qty"
        excelDoc.Cells(STARTING_ROW, INT_COL_MODEL_QTY) = "Model Qty"
        excelDoc.Cells(STARTING_ROW, INT_COL_TOTAL_QTY) = "Total Qty"
        excelDoc.Cells(STARTING_ROW, INT_COL_MIN_QTY) = "Min Qty"
        excelDoc.Cells(STARTING_ROW, INT_COL_AVG_QTY) = "Avg Qty"
        excelDoc.Cells(STARTING_ROW, INT_COL_MAX_QTY) = "Max Qty"
    End Sub

    '------------------------------------------------------------
    '-             Subprogram Name: writeEmployees              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 04/05/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Writes all the employees and their data to the excel     -
    '- spreadsheet                                              -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- excelDoc - The Excel.Application object to be used       -
    '-            within the subroutine                         -
    '- employees - A List of Employees to write to the Excel    -
    '-             spreadsheet                                  -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- FIRST_ROW - The first row that an employee should be     -
    '-             written to                                   -
    '- empl - An Employee object that is used to store the      -
    '-        current Employee                                  -
    '- row - What row is being written to for the employee      -
    '------------------------------------------------------------
    Private Sub writeEmployees(ByRef excelDoc As Excel.Application, employees As List(Of Employee))
        Const FIRST_ROW As Integer = STARTING_ROW + 1
        Dim empl As Employee
        Dim row As Integer

        For idx As Integer = 0 To employees.Count - 1
            row = FIRST_ROW + idx
            empl = employees(idx)

            writeEmployee(row, excelDoc, empl)
            maxRowWithData = row
        Next
    End Sub

    '------------------------------------------------------------
    '-              Subprogram Name: writeEmployee              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 04/05/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Writes an employee and their data to the excel           -
    '- spreadsheet                                              -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- row - The row number to write the Employee to            -
    '- excelDoc - The Excel.Application to write the data to    -
    '- empl - The Employee object to get data from              -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub writeEmployee(row As Integer, ByRef excelDoc As Excel.Application, empl As Employee)
        excelDoc.Cells(row, INT_COL_FIRST_NAME) = empl.firstName
        excelDoc.Cells(row, INT_COL_LAST_NAME) = empl.lastName
        excelDoc.Cells(row, INT_COL_ORDER_ID) = empl.orderId
        excelDoc.Cells(row, INT_COL_EMP_ID) = empl.id

        excelDoc.Cells(row, INT_COL_GAME_SALES) = empl.gameSales
        excelDoc.Cells(row, INT_COL_DOLL_SALES) = empl.dollSales
        excelDoc.Cells(row, INT_COL_BUILD_SALES) = empl.buildingSales
        excelDoc.Cells(row, INT_COL_MODEL_SALES) = empl.modelSales

        excelDoc.Cells(row, INT_COL_TOTAL_SALES) = String.Format(FORMULA_TOTAL_SALES, row)
        excelDoc.Cells(row, INT_COL_MIN_SALES) = String.Format(FORMULA_MIN_SALES, row)
        excelDoc.Cells(row, INT_COL_AVG_SALES) = String.Format(FORMULA_AVG_SALES, row)
        excelDoc.Cells(row, INT_COL_MAX_SALES) = String.Format(FORMULA_MAX_SALES, row)

        excelDoc.Cells(row, INT_COL_GAMES_QTY) = empl.gameQuantity
        excelDoc.Cells(row, INT_COL_DOLLS_QTY) = empl.dollQuantity
        excelDoc.Cells(row, INT_COL_BUILD_QTY) = empl.buildingQuanity
        excelDoc.Cells(row, INT_COL_MODEL_QTY) = empl.modelQuantity

        excelDoc.Cells(row, INT_COL_TOTAL_QTY) = String.Format(FORMULA_TOTAL_QTY, row)
        excelDoc.Cells(row, INT_COL_MIN_QTY) = String.Format(FORMULA_MIN_QTY, row)
        excelDoc.Cells(row, INT_COL_AVG_QTY) = String.Format(FORMULA_AVG_QTY, row)
        excelDoc.Cells(row, INT_COL_MAX_QTY) = String.Format(FORMULA_MAX_QTY, row)
    End Sub

    '------------------------------------------------------------
    '-           Subprogram Name: writeAggregateRows            -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 04/05/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Writes all of the aggregate data rows about all the      -
    '- employees.                                               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- excelDoc - The Excel.Application to write to             -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- FIRST_DATA_ROW - The first row with employee data        -
    '- FIRST_ROW - The first row to write to                    -
    '- formula - Current formula                                -
    '- formulaRow - Current row being written to                -
    '- functions - An array of strings containing various excel -
    '-             functions to be used                         -
    '- names - Function names to be used as titles for the      -
    '-         aggregate rows                                   -
    '------------------------------------------------------------
    Private Sub writeAggregateRows(ByRef excelDoc As Excel.Application)
        Dim FIRST_ROW As Integer = maxRowWithData + 2  ' 1 for next line, 1 for blank line
        Dim FIRST_DATA_ROW As Integer = STARTING_ROW + 1
        Dim functions As String() = {"sum", "max", "average", "min"}
        Dim names As String() = {"Total", "Max", "Avg", "Min"}
        Dim formula As String
        Dim formulaRow As Integer

        For idx As Integer = 0 To functions.Length - 1
            formulaRow = FIRST_ROW + idx
            formula = String.Format("=" & functions(idx) & FORMULA_AGGREGATE, "{0}", FIRST_DATA_ROW, maxRowWithData)

            excelDoc.Cells(formulaRow, INT_COL_AGG_TITLES) = names(idx) & ":"
            excelDoc.Cells(formulaRow, INT_COL_GAME_SALES) = String.Format(formula, COL_GAME_SALES)
            excelDoc.Cells(formulaRow, INT_COL_DOLL_SALES) = String.Format(formula, COL_DOLL_SALES)
            excelDoc.Cells(formulaRow, INT_COL_BUILD_SALES) = String.Format(formula, COL_BUILD_SALES)
            excelDoc.Cells(formulaRow, INT_COL_MODEL_SALES) = String.Format(formula, COL_MODEL_SALES)
            excelDoc.Cells(formulaRow, INT_COL_TOTAL_SALES) = String.Format(formula, COL_TOTAL_SALES)
            excelDoc.Cells(formulaRow, INT_COL_MIN_SALES) = String.Format(formula, COL_MIN_SALES)
            excelDoc.Cells(formulaRow, INT_COL_AVG_SALES) = String.Format(formula, COL_AVG_SALES)
            excelDoc.Cells(formulaRow, INT_COL_MAX_SALES) = String.Format(formula, COL_MAX_SALES)

            excelDoc.Cells(formulaRow, INT_COL_GAMES_QTY) = String.Format(formula, COL_GAMES_QTY)
            excelDoc.Cells(formulaRow, INT_COL_DOLLS_QTY) = String.Format(formula, COL_DOLLS_QTY)
            excelDoc.Cells(formulaRow, INT_COL_BUILD_QTY) = String.Format(formula, COL_BUILD_QTY)
            excelDoc.Cells(formulaRow, INT_COL_MODEL_QTY) = String.Format(formula, COL_MODEL_QTY)
            excelDoc.Cells(formulaRow, INT_COL_TOTAL_QTY) = String.Format(formula, COL_TOTAL_QTY)
            excelDoc.Cells(formulaRow, INT_COL_MIN_QTY) = String.Format(formula, COL_MIN_QTY)
            excelDoc.Cells(formulaRow, INT_COL_AVG_QTY) = String.Format(formula, COL_AVG_QTY)
            excelDoc.Cells(formulaRow, INT_COL_MAX_QTY) = String.Format(formula, COL_MAX_QTY)
        Next
    End Sub

    '------------------------------------------------------------
    '-         Function Name: ColumnLetterToColumnIndex         -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 04/05/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Convert a letter to an integer to be used with an        -
    '- Excel.Application's Cells property.                      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- columnLetter - The column letter to get the index for    -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- charA - The integer representation of the capital letter -
    '-         A                                                -
    '- charColLetter - The integer representation of the        -
    '-                 current letter in the columnLetter       -
    '- sum - The sum of addition and mulitplication for the     -
    '-       letter                                             -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Integer - The index that can be used with                -
    '-           Excel.Application.Cells                        -
    '------------------------------------------------------------
    Private Shared Function ColumnLetterToColumnIndex(columnLetter As String) As Integer
        ' Inspired by: https://www.add-in-express.com/creating-addins-blog/2013/11/13/convert-excel-column-number-to-name/
        columnLetter = columnLetter.ToUpper()
        Dim sum As Integer = 0
        Dim charA As Integer = Asc("A")
        Dim charColLetter As Integer

        For i As Integer = 0 To columnLetter.Length - 1
            sum *= 26
            charColLetter = Asc(columnLetter(i))
            sum += (charColLetter - charA) + 1
        Next

        Return sum
    End Function
End Class
