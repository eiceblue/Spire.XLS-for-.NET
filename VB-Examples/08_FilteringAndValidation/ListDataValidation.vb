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
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "DataValidation.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\DataValidation.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text for cells A7 to A10 with different city names
            sheet.Range("A7").Text = "Beijing"
            sheet.Range("A8").Text = "New York"
            sheet.Range("A9").Text = "Denver"
            sheet.Range("A10").Text = "Paris"

            ' Specify the range of cell D10 for data validation
            Dim range As CellRange = sheet.Range("D10")

            ' Enable error display for the data validation
            range.DataValidation.ShowError = True

            ' Set the alert style to stop, indicating that an error should stop the input
            range.DataValidation.AlertStyle = AlertStyleType.Stop

            ' Set the error title and message for the data validation
            range.DataValidation.ErrorTitle = "Error"
            range.DataValidation.ErrorMessage = "Please select a city from the list"

            ' Set the data range for the data validation to be cells A7 to A10
            range.DataValidation.DataRange = sheet.Range("A7:A10")

            ' Specify the output filename for the modified workbook
            Dim output As String = "ListDataValidation_out.xlsx"

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
