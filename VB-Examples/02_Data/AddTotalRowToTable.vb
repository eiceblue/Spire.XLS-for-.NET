Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core

Namespace AddTotalRowToTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AddATotalRowToTable.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create a table with the data from the specific cell range.
			Dim table As IListObject = sheet.ListObjects.Create("Table", sheet.Range("A1:D4"))

			'Display total row.
			table.DisplayTotalRow = True

			'Add a total row.
			table.Columns(0).TotalsRowLabel = "Total"
			table.Columns(1).TotalsCalculation = ExcelTotalsCalculation.Sum
			table.Columns(2).TotalsCalculation = ExcelTotalsCalculation.Sum
			table.Columns(3).TotalsCalculation = ExcelTotalsCalculation.Sum

			Dim result As String = "Result-AddATotalRowToTable.xlsx"

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
