'------------------------------------------------------------
'-                  File Name: Employee.vb                  -
'-                 Part of Project: Assign5                 -
'------------------------------------------------------------
'-                Written By: Elijah Wilson                 -
'-                  Written On: 02/13/2016                  -
'------------------------------------------------------------
'- File Purpose:                                            -
'-                                                          -
'- This contains the Employee class which is used to store  -
'- information about an employee and their sales.           -
'------------------------------------------------------------
Public Class Employee
    Public id As Integer
    Public firstName As String
    Public lastName As String
    Public orderId As Integer

    Public gameSales As Single
    Public gameQuantity As Integer

    Public dollSales As Single
    Public dollQuantity As Integer

    Public buildingSales As Single
    Public buildingQuanity As Integer

    Public modelSales As Single
    Public modelQuantity As Integer

    Public totalSales As Single
End Class