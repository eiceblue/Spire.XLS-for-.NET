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
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "DataValidation.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\DataValidation.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text for cell C12 on the worksheet
            sheet.Range("C12").Text = "Please enter number between 10 and 100:"

            ' Autofit the width of the columns in the range C12
            sheet.Range("C12").AutoFitColumns()

            ' Specify the range of cell D12 and assign it to the variable "range"
            Dim range As CellRange = sheet.Range("D12")

            ' Set the data validation to allow values of type Integer
            range.DataValidation.AllowType = CellDataType.Integer

            ' Set the data validation comparison operator to Between
            range.DataValidation.CompareOperator = ValidationComparisonOperator.Between

            ' Set the lower and upper limits for the number range in the data validation
            range.DataValidation.Formula1 = "10"
            range.DataValidation.Formula2 = "100"

            ' Set the alert style to Info, indicating that an information icon should be displayed for invalid entries
            range.DataValidation.AlertStyle = AlertStyleType.Info

            ' Enable error display for the data validation
            range.DataValidation.ShowError = True

            ' Set the error title and message for the data validation
            range.DataValidation.ErrorTitle = "Error"
            range.DataValidation.ErrorMessage = "Please enter a valid number"

            ' Set the input message for the data validation
            range.DataValidation.InputMessage = "Whole Number Validation Type"

            ' Ignore blank cells for the data validation
            range.DataValidation.IgnoreBlank = True

            ' Show input message for the data validation
            range.DataValidation.ShowInput = True

            ' Specify the output filename for the modified workbook
            Dim output As String = "WholeNumberDataValidation_out.xlsx"

            ' Save the modified workbook to a new Excel file with Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
