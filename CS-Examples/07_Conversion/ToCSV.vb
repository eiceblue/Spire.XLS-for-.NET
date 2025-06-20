Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ToCSV
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ToCSV.xlsx")

            ' Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Save the worksheet to a CSV file with comma as the delimiter and UTF-8 encoding.
            sheet.SaveToFile("ToCSV.csv", ",", Encoding.UTF8)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'view the document
            ExcelDocViewer("ToCSV.csv")
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
