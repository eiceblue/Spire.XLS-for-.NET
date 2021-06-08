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
			'Create a workbook
			Dim workbook As New Workbook()

			'Load a csv file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CSVToExcel.csv", ",", 1, 1)

			Dim sheet As Worksheet = workbook.Worksheets(0)
			sheet.Range("D2:E19").IgnoreErrorOptions = IgnoreErrorType.NumberAsText
			sheet.AllocatedRange.AutoFitColumns()

			'Save the document and launch it
			workbook.SaveToFile("CSVToExcel_result.xlsx", ExcelVersion.Version2013)
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
