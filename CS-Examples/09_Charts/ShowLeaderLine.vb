Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace ShowLeaderLine
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook
            Dim workbook As New Workbook()
            workbook.Version = ExcelVersion.Version2013

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set values in specific cells of the worksheet
            sheet.Range("A1").Value = "1"
            sheet.Range("A2").Value = "2"
            sheet.Range("A3").Value = "3"
            sheet.Range("B1").Value = "4"
            sheet.Range("B2").Value = "5"
            sheet.Range("B3").Value = "6"
            sheet.Range("C1").Value = "7"
            sheet.Range("C2").Value = "8"
            sheet.Range("C3").Value = "9"

            ' Add a new chart to the worksheet
            Dim chart As Chart = sheet.Charts.Add(ExcelChartType.BarStacked)
            chart.DataRange = sheet.Range("A1:C3")
            chart.TopRow = 4
            chart.LeftColumn = 2
            chart.Width = 450
            chart.Height = 300

            ' Configure data labels and leader lines for each series in the chart
            For Each cs As ChartSerie In chart.Series
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
                cs.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = True
            Next cs

            ' Save the workbook to a file named "Output.xlsx" in Excel 2013 format
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()
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
