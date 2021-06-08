Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace DataCallout
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Workbook
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\DataCallout.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the first chart
			Dim chart As Chart = sheet.Charts(0)

			For Each cs As ChartSerie In chart.Series
				cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
				cs.DataPoints.DefaultDataPoint.DataLabels.HasWedgeCallout = True
				cs.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = True
				cs.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = True
				cs.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = True
			Next cs

			'Save and Launch
			workbook.SaveToFile("Output.xlsx", FileFormat.Version2010)
			ExcelDocViewer(workbook.FileName)
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
