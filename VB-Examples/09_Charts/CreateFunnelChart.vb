Imports Spire.Xls

Namespace CreateFunnelChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "Funnel.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Funnel.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet = workbook.Worksheets(0)

            ' Add a chart to the worksheet
            Dim officeChart = sheet.Charts.Add()

            ' Set the chart type to Funnel
            officeChart.ChartType = ExcelChartType.Funnel

            ' Set the data range for the chart
            officeChart.DataRange = sheet.Range("A1:B6")

            ' Set the chart title to "Funnel"
            officeChart.ChartTitle = "Funnel"

            ' Disable the legend in the chart
            officeChart.HasLegend = False

            ' Enable data labels with values for each data point in the chart series
            officeChart.Series(0).DataPoints.DefaultDataPoint.DataLabels.HasValue = True

            ' Set the font size of the data labels to 8
            officeChart.Series(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8

            ' Save the modified workbook to "Funnel_chart.xlsx"
            workbook.SaveToFile("Funnel_chart.xlsx")
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("Funnel_chart.xlsx")
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
