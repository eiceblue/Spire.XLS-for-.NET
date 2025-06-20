Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace CreateSunBurstChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "SunBurst.xlsx"
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SunBurst.xlsx")

            ' Get the first worksheet from the loaded workbook
            Dim sheet = workbook.Worksheets(0)

            ' Add a chart to the worksheet
            Dim officeChart = sheet.Charts.Add()

            ' Set the chart type to Sunburst
            officeChart.ChartType = ExcelChartType.SunBurst

            ' Specify the data range for the chart
            officeChart.DataRange = sheet("A1:D16")

            ' Set the top row, bottom row, left column, and right column of the chart
            officeChart.TopRow = 1
            officeChart.BottomRow = 17
            officeChart.LeftColumn = 6
            officeChart.RightColumn = 14

            ' Set the chart title
            officeChart.ChartTitle = "Sales by quarter"

            ' Set the size of data labels for the series' default data points
            officeChart.Series(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8

            ' Disable the chart legend
            officeChart.HasLegend = False

            ' Save the workbook with the chart to a new file named "Sunburst_chart.xlsx"
            workbook.SaveToFile("Sunburst_chart.xlsx")
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("Sunburst_chart.xlsx")
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
