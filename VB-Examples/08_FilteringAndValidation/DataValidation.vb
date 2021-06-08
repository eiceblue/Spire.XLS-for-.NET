Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace DataValidation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\DataValidation.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Decimal DataValidation
			sheet.Range("B11").Text = "Input Number(3-6):"
			Dim rangeNumber As CellRange = sheet.Range("B12")
			'Set the operator for the data validation.
			rangeNumber.DataValidation.CompareOperator = ValidationComparisonOperator.Between
			'Set the value or expression associated with the data validation.
			rangeNumber.DataValidation.Formula1 = "3"
			'The value or expression associated with the second part of the data validation.
			rangeNumber.DataValidation.Formula2 = "6"
			'Set the data validation type.
			rangeNumber.DataValidation.AllowType = CellDataType.Decimal
			'Set the data validation error message.
			rangeNumber.DataValidation.ErrorMessage = "Please input correct number!"
			'Enable the error.
			rangeNumber.DataValidation.ShowError = True
			rangeNumber.Style.KnownColor = ExcelColors.Gray25Percent

			'Date DataValidation
			sheet.Range("B14").Text = "Input Date:"
			Dim rangeDate As CellRange = sheet.Range("B15")
			rangeDate.DataValidation.AllowType = CellDataType.Date
			rangeDate.DataValidation.CompareOperator = ValidationComparisonOperator.Between
			rangeDate.DataValidation.Formula1= "1/1/1970"
			rangeDate.DataValidation.Formula2 = "12/31/1970"
			rangeDate.DataValidation.ErrorMessage = "Please input correct date!"
			rangeDate.DataValidation.ShowError = True
			rangeDate.DataValidation.AlertStyle = AlertStyleType.Warning
			rangeDate.Style.KnownColor = ExcelColors.Gray25Percent

			'TextLength DataValidation
			sheet.Range("B17").Text = "Input Text:"
			Dim rangeTextLength As CellRange = sheet.Range("B18")
			rangeTextLength.DataValidation.AllowType = CellDataType.TextLength
			rangeTextLength.DataValidation.CompareOperator = ValidationComparisonOperator.LessOrEqual
			rangeTextLength.DataValidation.Formula1 = "5"
			rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!"
			rangeTextLength.DataValidation.ShowError = True
			rangeTextLength.DataValidation.AlertStyle = AlertStyleType.Stop
			rangeTextLength.Style.KnownColor = ExcelColors.Gray25Percent

			sheet.AutoFitColumn(2)

			Dim result As String="DataValidation_result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
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
