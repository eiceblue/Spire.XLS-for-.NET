Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core

Namespace ListDataValidation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\DataValidation.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set text for cells 
			sheet.Range("A7").Text = "Beijing"
			sheet.Range("A8").Text = "New York"
			sheet.Range("A9").Text = "Denver"
			sheet.Range("A10").Text = "Paris"

			'Set data validation for cell
			Dim range As CellRange = sheet.Range("D10")
			range.DataValidation.ShowError = True
			range.DataValidation.AlertStyle = AlertStyleType.Stop
			range.DataValidation.ErrorTitle = "Error"
			range.DataValidation.ErrorMessage = "Please select a city from the list"
			range.DataValidation.DataRange = sheet.Range("A7:A10")

			'Save the document
			Dim output As String = "ListDataValidation_out.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(output)
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
