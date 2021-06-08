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
			'Create a workbook.
			Dim workbook As New Workbook()

			'Create an empty worksheet.
			workbook.CreateEmptySheets(1)

			'Get the worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the value of max row and column.
			Dim maxRow As Integer = 10000
			'int maxRow = 5;
			Dim maxCol As Integer = 200
			'int maxCol = 5;

			'Output an array of data to a range of worksheet.
			Dim myarray(maxRow, maxCol) As Object
			Dim isred(maxRow, maxCol) As Boolean
			For i As Integer = 0 To maxRow
				For j As Integer = 0 To maxCol
					myarray(i, j) = i + j
					If CInt(Fix(myarray(i, j))) > 8 Then
						isred(i, j) = True
					End If
				Next j
			Next i

			sheet.InsertArray(myarray, 1, 1)

			Dim result As String = "Result-SetArrayOfValuesIntoRange.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

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
