Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace SetArrayOfValuesIntoRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Creates a new Excel workbook.
            Dim workbook As New Workbook()

            ' Creates a single empty worksheet in the workbook.
            workbook.CreateEmptySheets(1)

            ' Retrieves the first worksheet in the workbook (index starts at 0).
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Specifies the maximum number of rows.
            Dim maxRow As Integer = 10000
            ' Specifies the maximum number of columns.
            Dim maxCol As Integer = 200

            ' Declares a 2D array to store data.
            Dim myarray(maxRow, maxCol) As Object
            ' Declares a 2D array to indicate if a cell should be red.
            Dim isred(maxRow, maxCol) As Boolean
            ' Iterates through each row.
            For i As Integer = 0 To maxRow
                ' Iterates through each column.
                For j As Integer = 0 To maxCol
                    ' Computes the value to be assigned to the current cell based on its row and column index.
                    myarray(i, j) = i + j
                    ' Checks if the computed value is greater than 8.
                    If CInt(Fix(myarray(i, j))) > 8 Then
                        ' Sets the corresponding flag to indicate that the cell should be red.
                        isred(i, j) = True
                    End If
                Next j
            Next i
            ' Inserts the array of data into the worksheet starting from cell A1.
            sheet.InsertArray(myarray, 1, 1)
            ' Defines the file name for the resulting Excel file.
            Dim result As String = "Result-SetArrayOfValuesIntoRange.xlsx"

            ' Saves the workbook to a file with the specified file name in the Excel 2013 format.
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
            ExcelDocViewer(result)
		End Sub

		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
