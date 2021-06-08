Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ToCSVWithFilteredValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AutofilterSample.xlsx")

			'Convert to CSV file with filtered value
			workbook.Worksheets(0).SaveToFile("ToCSVWithFilteredValue.csv", ";", False)

			'Convert to CSV stream
			'worksheet.SaveToStream(Stream stream, string separator, bool retainHiddenData);           

			'View the document
			FileViewer("ToCSVWithFilteredValue.csv")
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
