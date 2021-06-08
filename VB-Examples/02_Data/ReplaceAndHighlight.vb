Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ReplaceAndHighlight
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ReplaceAndHighlight.xlsx")

			Dim worksheet As Worksheet = workbook.Worksheets(0)

			Dim ranges() As CellRange = worksheet.FindAllString("Total", True, True)

			For Each range As CellRange In ranges
				'reset the text, in other words, replace the text
				range.Text = "Sum"

				'set the color
				range.Style.Color = Color.Yellow
			Next range

			Dim result As String="ReplaceAndHighlight_result.xlsx"
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
