Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.ComponentModel
Imports System.Text

Namespace ChangeSeriesColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChangeSeriesColor.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the first chart
			Dim chart As Chart = sheet.Charts(0)

			'Get the second series
			Dim cs As ChartSerie = chart.Series(1)

			'Set the fill type
			cs.Format.Fill.FillType = ShapeFillType.SolidColor

			'Change the fill color
			cs.Format.Fill.ForeColor = Color.Orange

			'Save and Launch
			workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
			ExcelDocViewer("Output.xlsx")
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
