Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace CreateHistogramChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "HistogramChart.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\HistogramChart.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet = workbook.Worksheets(0)

            ' Add a chart to the worksheet
            Dim officeChart = sheet.Charts.Add()

            ' Set the chart type to Histogram
            officeChart.ChartType = ExcelChartType.Histogram

            ' Set the data range for the chart
            officeChart.DataRange = sheet("A1:A15")

            ' Set the position of the chart within the worksheet using row and column indices
            officeChart.TopRow = 1
            officeChart.BottomRow = 19
            officeChart.LeftColumn = 4
            officeChart.RightColumn = 12

            ' Set the bin width for the primary category axis (X-axis)
            officeChart.PrimaryCategoryAxis.BinWidth = 8

            ' Set the gap width between bars in the chart series
            officeChart.Series(0).DataFormat.Options.GapWidth = 6

            ' Set the chart title to "Height Data"
            officeChart.ChartTitle = "Height Data"

            ' Set the title for the primary value axis (Y-axis)
            officeChart.PrimaryValueAxis.Title = "Number of students"

            ' Set the title for the primary category axis (X-axis)
            officeChart.PrimaryCategoryAxis.Title = "Height"

            ' Disable the legend in the chart
            officeChart.HasLegend = False

            ' Save the modified workbook to "Histogram_chart.xlsx"
            workbook.SaveToFile("Histogram_chart.xlsx")
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("Histogram_chart.xlsx")
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
