Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace SetFont
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a Workbook from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SetFont.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the first sheet
			Dim chart As Chart = sheet.Charts(0)

			'Create a font
			Dim font As ExcelFont = workbook.CreateFont()
			font.Size = 15.0
			font.Color = Color.LightSeaGreen

			For Each cs As ChartSerie In chart.Series
				'Set font
				cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font)
			Next cs

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
