Imports Spire.Xls

Namespace CreateTreeMapChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "TreeMap.xlsx"
            workbook.LoadFromFile("..\..\..\..\..\..\Data\TreeMap.xlsx")

            ' Get the first worksheet from the loaded workbook
            Dim sheet = workbook.Worksheets(0)

            ' Add a chart to the worksheet
            Dim officeChart = sheet.Charts.Add()

            ' Set the chart type to TreeMap
            officeChart.ChartType = ExcelChartType.TreeMap

            ' Specify the data range for the chart
            officeChart.DataRange = sheet("A2:C11")

            ' Set the top row, bottom row, left column, and right column of the chart
            officeChart.TopRow = 1
            officeChart.BottomRow = 19
            officeChart.LeftColumn = 4
            officeChart.RightColumn = 14

            ' Set the chart title
            officeChart.ChartTitle = "Area by countries"

            ' Set the label option for the TreeMap series to use banners
            officeChart.Series(0).DataFormat.TreeMapLabelOption = ExcelTreeMapLabelOption.Banner

            ' Set the size of data labels for the series' default data points
            officeChart.Series(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8

            ' Save the workbook with the chart to a new file named "treemap_chart.xlsx"
            workbook.SaveToFile("treemap_chart.xlsx")
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("treemap_chart.xlsx")
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
