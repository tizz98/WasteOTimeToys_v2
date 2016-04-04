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
Imports System.Reflection

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

    '------------------------------------------------------------
    '-                 Function Name: toString                  -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a string representation of the Employee object.  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- accStr - A string that is accumulated with data based on -
    '-          the fields                                      -
    '- fields - An array of FieldInfo objects, based off of     -
    '-          this class's type                               -
    '- fmtStr - A string that will be formatted with the accStr -
    '-          and the class's type's FullName                 -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - A string representation of the Employee object  -
    '------------------------------------------------------------
    Public Overrides Function toString() As String
        Dim fields As FieldInfo() = Me.GetType().GetFields()
        Dim accStr As String = ""
        Dim fmtStr As String = "<" & Me.GetType().FullName & "({0})>"

        For Each field As FieldInfo In fields
            If Not field.IsSpecialName Then
                accStr &= String.Format("{0}: {1}, ", field.Name, field.GetValue(Me))
            End If
        Next

        accStr = accStr.TrimEnd(" ").TrimEnd(",")

        Return String.Format(fmtStr, accStr)
    End Function

    '------------------------------------------------------------
    '-                 Function Name: fullName                  -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- The Employee's last name comma first name.               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - A string concatenation of the Employee's last   -
    '-          name and firstname                              -
    '------------------------------------------------------------
    Public Function fullName() As String
        Return String.Format("{0}, {1}", Me.lastName, Me.firstName)
    End Function
End Class