Imports Microsoft.Office.Interop

Public Class frmToys
    Private Sub btnMagic_Click(sender As Object, e As EventArgs) Handles btnMagic.Click
        Dim checkExcel As Object
        Dim excelDoc As Excel.Application

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
        excelDoc.Visible = True

        excelDoc.Workbooks.Add()
        excelDoc.Sheets.Add()

        ' example
        ' excelDoc.Cells(row, col) = 123
        ' excelDoc.Cells(row, col) = '=average(a1...a5)'

        MessageBox.Show("Pausing...")
        excelDoc.Quit()
        excelDoc = Nothing
    End Sub
End Class
