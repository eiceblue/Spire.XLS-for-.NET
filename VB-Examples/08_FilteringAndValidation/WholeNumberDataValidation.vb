Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core

Namespace WholeNumberDataValidation
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

			sheet.Range("C12").Text = "Please enter number between 10 and 100:"
			sheet.Range("C12").AutoFitColumns()

			'Set Whole Number data validation for cell "D12"
			Dim range As CellRange = sheet.Range("D12")
			range.DataValidation.AllowType = CellDataType.Integer
			range.DataValidation.CompareOperator = ValidationComparisonOperator.Between

			range.DataValidation.Formula1 = "10"
			range.DataValidation.Formula2 = "100"

			range.DataValidation.AlertStyle = AlertStyleType.Info
			range.DataValidation.ShowError = True
			range.DataValidation.ErrorTitle = "Error"
			range.DataValidation.ErrorMessage = "Please enter a valid number"
			range.DataValidation.InputMessage = "Whole Number Validation Type"
			range.DataValidation.IgnoreBlank = True
			range.DataValidation.ShowInput = True

			'Save the document
			Dim output As String = "WholeNumberDataValidation_out.xlsx"
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
