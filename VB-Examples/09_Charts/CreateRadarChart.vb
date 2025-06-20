Imports Spire.Xls

Namespace CreateRadarChart

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Create one empty worksheet in the workbook
            workbook.CreateEmptySheets(1)

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the name of the worksheet to "Chart data"
            sheet.Name = "Chart data"

            ' Hide gridlines on the worksheet
            sheet.GridLinesVisible = False

            ' Call a function to create chart data in the worksheet
            CreateChartData(sheet)

            ' Add a chart to the worksheet
            Dim chart As Chart = sheet.Charts.Add()

            ' Set the position of the chart within the worksheet using column and row indices
            chart.LeftColumn = 1
            chart.TopRow = 6
            chart.RightColumn = 11
            chart.BottomRow = 29

            ' Set the data range for the chart
            chart.DataRange = sheet.Range("A1:C5")

            ' Specify that series data will not be derived from the range itself
            chart.SeriesDataFromRange = False

            ' Check if checkbox1 is checked
            If checkBox1.Checked Then
                ' Set the chart type to filled radar
                chart.ChartType = ExcelChartType.RadarFilled
            Else
                ' Set the chart type to radar
                chart.ChartType = ExcelChartType.Radar
            End If

            ' Set the chart title to "Sale market by region"
            chart.ChartTitle = "Sale market by region"

            ' Make the chart title bold and set the font size to 12
            chart.ChartTitleArea.IsBold = True
            chart.ChartTitleArea.Size = 12

            ' Hide the fill in the plot area of the chart
            chart.PlotArea.Fill.Visible = False

            ' Set the position of the legend to the corner of the chart
            chart.Legend.Position = LegendPositionType.Corner

            ' Specify the output file name as "CreateRadarChart.xlsx"
            Dim output As String = "CreateRadarChart.xlsx"

            ' Save the modified workbook to the specified file path, using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
            ExcelDocViewer(output)
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
            ' Set the value "Product" in cell A1
            sheet.Range("A1").Value = "Product"

            ' Set the values for different products in column A
            sheet.Range("A2").Value = "Bikes"
            sheet.Range("A3").Value = "Cars"
            sheet.Range("A4").Value = "Trucks"
            sheet.Range("A5").Value = "Buses"

            ' Set the value "Paris" in cell B1
            sheet.Range("B1").Value = "Paris"

            ' Set the numeric values for Paris sales in column B
            sheet.Range("B2").NumberValue = 4000
            sheet.Range("B3").NumberValue = 23000
            sheet.Range("B4").NumberValue = 4000
            sheet.Range("B5").NumberValue = 30000

            ' Set the value "New York" in cell C1
            sheet.Range("C1").Value = "New York"

            ' Set the numeric values for New York sales in column C
            sheet.Range("C2").NumberValue = 30000
            sheet.Range("C3").NumberValue = 7600
            sheet.Range("C4").NumberValue = 18000
            sheet.Range("C5").NumberValue = 8000

            ' Apply formatting to the ranges

            ' Make cells A1 to C1 bold
            sheet.Range("A1:C1").Style.Font.IsBold = True

            ' Set a light yellow background color for cells A2 to C2
            sheet.Range("A2:C2").Style.KnownColor = ExcelColors.LightYellow

            ' Set a light green background color for cells A3 to C3
            sheet.Range("A3:C3").Style.KnownColor = ExcelColors.LightGreen1

            ' Set a light orange background color for cells A4 to C4
            sheet.Range("A4:C4").Style.KnownColor = ExcelColors.LightOrange

            ' Set a light turquoise background color for cells A5 to C5
            sheet.Range("A5:C5").Style.KnownColor = ExcelColors.LightTurquoise

            ' Apply border formatting to the range A1 to C5

            ' Set a thin border with blue color for the top edge
            sheet.Range("A1:C5").Style.Borders(BordersLineType.EdgeTop).Color = Color.FromArgb(0, 0, 128)
            sheet.Range("A1:C5").Style.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin

            ' Set a thin border with blue color for the bottom edge
            sheet.Range("A1:C5").Style.Borders(BordersLineType.EdgeBottom).Color = Color.FromArgb(0, 0, 128)
            sheet.Range("A1:C5").Style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Thin

            ' Set a thin border with blue color for the left edge
            sheet.Range("A1:C5").Style.Borders(BordersLineType.EdgeLeft).Color = Color.FromArgb(0, 0, 128)
            sheet.Range("A1:C5").Style.Borders(BordersLineType.EdgeLeft).LineStyle = LineStyleType.Thin

            ' Set a thin border with blue color for the right edge
            sheet.Range("A1:C5").Style.Borders(BordersLineType.EdgeRight).Color = Color.FromArgb(0, 0, 128)
            sheet.Range("A1:C5").Style.Borders(BordersLineType.EdgeRight).LineStyle = LineStyleType.Thin

            ' Set the number format of cells B2 to C5 as currency with thousands separator
            sheet.Range("B2:C5").Style.NumberFormat = """$""#,##0"
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
