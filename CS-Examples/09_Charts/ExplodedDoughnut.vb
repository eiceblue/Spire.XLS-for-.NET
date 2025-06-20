Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace ExplodedDoughnut
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Declare a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook and assign it to the sheet variable
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the name of the sheet as "ExplodedDoughnut"
            sheet.Name = "ExplodedDoughnut"

            ' Call a method to create chart data in the sheet
            CreateChartData(sheet)

            ' Add a new chart to the sheet
            Dim chart As Chart = sheet.Charts.Add()

            ' Set the chart type as DoughnutExploded
            chart.ChartType = ExcelChartType.DoughnutExploded

            ' Set the position and size of the chart on the sheet
            chart.LeftColumn = 1
            chart.TopRow = 6
            chart.RightColumn = 11
            chart.BottomRow = 29

            ' Set the data range for the chart
            chart.DataRange = sheet.Range("A1:B5")

            ' Specify that series data is not derived from the range
            chart.SeriesDataFromRange = False

            ' Set the title of the chart as "Sales market by country"
            chart.ChartTitle = "Sales market by country"

            ' Customize the appearance of the chart title
            chart.ChartTitleArea.IsBold = True
            chart.ChartTitleArea.Size = 12

            ' Iterate through each series in the chart
            For Each cs As ChartSerie In chart.Series
                ' Enable varying colors for each series
                cs.Format.Options.IsVaryColor = True
                ' Show data labels for each data point
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = True
            Next cs

            ' Hide the fill of the plot area
            chart.PlotArea.Fill.Visible = False

            ' Set the position of the legend at the top of the chart
            chart.Legend.Position = LegendPositionType.Top

            ' Save the workbook to a file named "Sample.xlsx" in Excel 2010 format
            workbook.SaveToFile("Sample.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Sample.xlsx")
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
            ' Set the values for the "Country" column
            sheet.Range("A1").Value = "Country"
            sheet.Range("A2").Value = "Cuba"
            sheet.Range("A3").Value = "Mexico"
            sheet.Range("A4").Value = "France"
            sheet.Range("A5").Value = "Germany"

            ' Set the values for the "Sales" column
            sheet.Range("B1").Value = "Sales"
            sheet.Range("B2").NumberValue = 6000
            sheet.Range("B3").NumberValue = 8000
            sheet.Range("B4").NumberValue = 9000
            sheet.Range("B5").NumberValue = 8500

            ' Customize the appearance of the header row (A1:B1)
            sheet.Range("A1:B1").RowHeight = 15
            sheet.Range("A1:B1").Style.Color = Color.DarkGray
            sheet.Range("A1:B1").Style.Font.Color = Color.White
            sheet.Range("A1:B1").Style.VerticalAlignment = VerticalAlignType.Center
            sheet.Range("A1:B1").Style.HorizontalAlignment = HorizontalAlignType.Center

            ' Apply number format to the cells in the "Sales" column (B2:B5)
            sheet.Range("B2:B5").Style.NumberFormat = """$""#,##0"
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
