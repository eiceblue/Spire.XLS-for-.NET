Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core

Namespace TimeDataValidation
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

			sheet.Range("C12").Text = "Please enter time between 09:00 and 18:00:"
			sheet.Range("C12").AutoFitColumns()

			'Set Time data validation for cell "D12"
			Dim range As CellRange = sheet.Range("D12")
			range.DataValidation.AllowType = CellDataType.Time
			range.DataValidation.CompareOperator = ValidationComparisonOperator.Between

			range.DataValidation.Formula1 = "09:00"
			range.DataValidation.Formula2 = "18:00"

			range.DataValidation.AlertStyle = AlertStyleType.Info
			range.DataValidation.ShowError = True
			range.DataValidation.ErrorTitle = "Time Error"
			range.DataValidation.ErrorMessage = "Please enter a valid time"
			range.DataValidation.InputMessage = "Time Validation Type"
			range.DataValidation.IgnoreBlank = True
			range.DataValidation.ShowInput = True

			'Save the document
			Dim output As String = "TimeDataValidation_out.xlsx"
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
