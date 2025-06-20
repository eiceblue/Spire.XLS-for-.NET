Imports Spire.Xls

Namespace CreateWaterfallChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "WaterfallChart.xlsx"
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WaterfallChart.xlsx")

            ' Get the first worksheet from the loaded workbook
            Dim sheet = workbook.Worksheets(0)

            ' Add a chart to the worksheet
            Dim officeChart = sheet.Charts.Add()

            ' Set the chart type to Waterfall
            officeChart.ChartType = ExcelChartType.WaterFall

            ' Specify the data range for the chart
            officeChart.DataRange = sheet("A2:B8")

            ' Set the top row, bottom row, left column, and right column of the chart
            officeChart.TopRow = 1
            officeChart.BottomRow = 19
            officeChart.LeftColumn = 4
            officeChart.RightColumn = 12

            ' Set the fourth and seventh data points of the series as total
            officeChart.Series(0).DataPoints(3).SetAsTotal = True
            officeChart.Series(0).DataPoints(6).SetAsTotal = True

            ' Show connector lines for the series
            officeChart.Series(0).Format.ShowConnectorLines = True

            ' Set the chart title
            officeChart.ChartTitle = "Waterfall Chart"

            ' Enable data labels for the default data points of the series
            officeChart.Series(0).DataPoints.DefaultDataPoint.DataLabels.HasValue = True
            officeChart.Series(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8

            ' Set the position of the legend to the right of the chart
            officeChart.Legend.Position = LegendPositionType.Right

            ' Save the workbook with the chart to a new file named "waterfall_chart.xlsx"
            workbook.SaveToFile("waterfall_chart.xlsx")
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("waterfall_chart.xlsx")
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
