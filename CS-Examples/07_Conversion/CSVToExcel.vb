Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CSVToExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the CSV file from the specified path with the specified delimiter (","),
            ' starting at row 1 and column 1
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CSVToExcel.csv", ",", 1, 1)

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the error options for the range D2:E19 to ignore number-as-text errors
            sheet.Range("D2:E19").IgnoreErrorOptions = IgnoreErrorType.NumberAsText

            ' Auto-fit the columns in the selected range to fit the content
            sheet.AllocatedRange.AutoFitColumns()

            ' Save the workbook to a new Excel file named "CSVToExcel_result.xlsx" in Excel 2013 format
            workbook.SaveToFile("CSVToExcel_result.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("CSVToExcel_result.xlsx")
		End Sub

		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
		Private Sub btnClose_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
