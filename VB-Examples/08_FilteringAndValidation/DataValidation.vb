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
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "DataValidation.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\DataValidation.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text for cell B11 to indicate the expected input for a number
            sheet.Range("B11").Text = "Input Number(3-6):"

            ' Define a range of cells (B12) for data validation related to numbers
            Dim rangeNumber As CellRange = sheet.Range("B12")
            rangeNumber.DataValidation.CompareOperator = ValidationComparisonOperator.Between
            rangeNumber.DataValidation.Formula1 = "3"
            rangeNumber.DataValidation.Formula2 = "6"
            rangeNumber.DataValidation.AllowType = CellDataType.Decimal
            rangeNumber.DataValidation.ErrorMessage = "Please input correct number!"
            rangeNumber.DataValidation.ShowError = True
            rangeNumber.Style.KnownColor = ExcelColors.Gray25Percent

            ' Set the text for cell B14 to indicate the expected input for a date
            sheet.Range("B14").Text = "Input Date:"

            ' Define a range of cells (B15) for data validation related to dates
            Dim rangeDate As CellRange = sheet.Range("B15")
            rangeDate.DataValidation.AllowType = CellDataType.Date
            rangeDate.DataValidation.CompareOperator = ValidationComparisonOperator.Between
            rangeDate.DataValidation.Formula1 = "1/1/1970"
            rangeDate.DataValidation.Formula2 = "12/31/1970"
            rangeDate.DataValidation.ErrorMessage = "Please input correct date!"
            rangeDate.DataValidation.ShowError = True
            rangeDate.DataValidation.AlertStyle = AlertStyleType.Warning
            rangeDate.Style.KnownColor = ExcelColors.Gray25Percent

            ' Set the text for cell B17 to indicate the expected input for a text
            sheet.Range("B17").Text = "Input Text:"

            ' Define a range of cells (B18) for data validation related to text length
            Dim rangeTextLength As CellRange = sheet.Range("B18")
            rangeTextLength.DataValidation.AllowType = CellDataType.TextLength
            rangeTextLength.DataValidation.CompareOperator = ValidationComparisonOperator.LessOrEqual
            rangeTextLength.DataValidation.Formula1 = "5"
            rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!"
            rangeTextLength.DataValidation.ShowError = True
            rangeTextLength.DataValidation.AlertStyle = AlertStyleType.Stop
            rangeTextLength.Style.KnownColor = ExcelColors.Gray25Percent

            ' Auto-fit column 2 (column B) to adjust the width based on content
            sheet.AutoFitColumn(2)

            ' Specify the output filename for the modified workbook
            Dim result As String = "DataValidation_result.xlsx"

            ' Save the modified workbook to a new Excel file with Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
