Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace RemoveDuplicatedRows
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new workbook
			Dim workbook As Workbook = New Workbook()

			' Load an existing workbook with a pivot table from a file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\DuplicatedRows.xlsx")

			' Get the first worksheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Remove duplicated rows in the worksheet
			sheet.RemoveDuplicates()

			' Remove the duplicate rows within the specified range
			' sheet.RemoveDuplicates(int startRow, int startColumn, int endRow, int endColumn)
			' Remove the duplicated rows based on specific columns and headers
			' sheet.RemoveDuplicates(int startRow, int startColumn, int endRow, int endColumn, boolean hasHeaders, int[] columnOffsets)

			' Specify the output file name for the result
			Dim result As String = "RemoveDuplicatedRows_result.xlsx"

			' Save the modified workbook to the specified file using Excel 2010 format
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			'View the document
			FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
