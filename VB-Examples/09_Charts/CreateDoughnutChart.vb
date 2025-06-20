Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CreateDoughnutChart

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the value in cell A1 as "Country"
            sheet.Range("A1").Value = "Country"

            ' Apply bold font style to cell A1
            sheet.Range("A1").Style.Font.IsBold = True

            ' Set the values in cells A2 to A5 for country names
            sheet.Range("A2").Value = "Cuba"
            sheet.Range("A3").Value = "Mexico"
            sheet.Range("A4").Value = "France"
            sheet.Range("A5").Value = "Germany"

            ' Set the value in cell B1 as "Sales"
            sheet.Range("B1").Value = "Sales"

            ' Apply bold font style to cell B1
            sheet.Range("B1").Style.Font.IsBold = True

            ' Set the numeric values in cells B2 to B5 for sales data
            sheet.Range("B2").NumberValue = 6000
            sheet.Range("B3").NumberValue = 8000
            sheet.Range("B4").NumberValue = 9000
            sheet.Range("B5").NumberValue = 8500

            ' Add a chart to the worksheet
            Dim chart As Chart = sheet.Charts.Add()

            ' Set the chart type to Doughnut
            chart.ChartType = ExcelChartType.Doughnut

            ' Set the data range for the chart
            chart.DataRange = sheet.Range("A1:B5")

            ' Specify that the series data will be manually specified (not derived from the range)
            chart.SeriesDataFromRange = False

            ' Set the position of the chart within the worksheet
            chart.LeftColumn = 4
            chart.TopRow = 2
            chart.RightColumn = 12
            chart.BottomRow = 22

            ' Set the chart title to "Market share by country"
            chart.ChartTitle = "Market share by country"

            ' Configure the chart title area formatting with bold text and font size 12
            chart.ChartTitleArea.IsBold = True
            chart.ChartTitleArea.Size = 12

            ' Enable data labels with percentage for each data point in the chart series
            For Each cs As ChartSerie In chart.Series
                cs.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = True
            Next cs

            ' Set the legend position to the top of the chart
            chart.Legend.Position = LegendPositionType.Top

            ' Specify the output file name as "CreateDoughnutChart.xlsx"
            Dim output As String = "CreateDoughnutChart.xlsx"

            ' Save the modified workbook to the specified file path, using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
            ExcelDocViewer(output)
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
