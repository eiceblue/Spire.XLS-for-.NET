Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace PyramidColumn
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

            ' Set the name of the worksheet to "Chart"
            sheet.Name = "Chart"

            ' Call the function to create chart data
            CreateChartData(sheet)

            ' Add a chart to the worksheet
            Dim chart As Chart = sheet.Charts.Add()

            ' Set the range of data for the chart
            chart.DataRange = sheet.Range("B2:B5")
            ' Disable automatic series data detection
            chart.SeriesDataFromRange = False
            ' Set the left column index for the chart position
            chart.LeftColumn = 1
            ' Set the top row index for the chart position
            chart.TopRow = 6
            ' Set the right column index for the chart position
            chart.RightColumn = 11
            ' Set the bottom row index for the chart position
            chart.BottomRow = 29

            If checkBox1.Checked Then
                ' Set the chart type to 3D clustered pyramid if the checkbox is checked
                chart.ChartType = ExcelChartType.Pyramid3DClustered
            Else
                ' Set the chart type to clustered pyramid if the checkbox is not checked
                chart.ChartType = ExcelChartType.PyramidClustered
            End If
            ' Set the title of the chart
            chart.ChartTitle = "Sales by year"
            ' Make the chart title text bold
            chart.ChartTitleArea.IsBold = True
            ' Set the font size of the chart title
            chart.ChartTitleArea.Size = 12
            ' Set the title of the primary category axis
            chart.PrimaryCategoryAxis.Title = "Year"
            ' Make the primary category axis title font bold
            chart.PrimaryCategoryAxis.Font.IsBold = True
            ' Make the primary category axis title area bold
            chart.PrimaryCategoryAxis.TitleArea.IsBold = True
            ' Set the title of the primary value axis
            chart.PrimaryValueAxis.Title = "Sales(in Dollars)"
            ' Disable major grid lines on the primary value axis
            chart.PrimaryValueAxis.HasMajorGridLines = False
            ' Set the minimum value of the primary value axis
            chart.PrimaryValueAxis.MinValue = 1000
            ' Make the primary value axis title area bold
            chart.PrimaryValueAxis.TitleArea.IsBold = True
            ' Set the rotation angle for the primary value axis title
            chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90
            ' Get the first series from the chart
            Dim cs As ChartSerie = chart.Series(0)
            ' Set the category labels for the series
            cs.CategoryLabels = sheet.Range("A2:A5")
            ' Enable varying colors for data points in the series
            cs.Format.Options.IsVaryColor = True
            ' Set the position of the chart legend to the top
            chart.Legend.Position = LegendPositionType.Top
            ' Save the workbook to a file named "Output.xlsx"
            workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer("Output.xlsx")
		End Sub

		Private Sub CreateChartData(ByVal sheet As Worksheet)
            'Set the value of cells A1 to B5
            sheet.Range("A1").Value = "Year"
            sheet.Range("A2").Value = "2002"
            sheet.Range("A3").Value = "2003"
            sheet.Range("A4").Value = "2004"
            sheet.Range("A5").Value = "2005"

            sheet.Range("B1").Value = "Sales"
            sheet.Range("B2").NumberValue = 4000
            sheet.Range("B3").NumberValue = 6000
            sheet.Range("B4").NumberValue = 7000
            sheet.Range("B5").NumberValue = 8500

            ' Set the row height of cells A1 to B1
            sheet.Range("A1:B1").RowHeight = 15
            ' Set the background color of cells A1 to B1
            sheet.Range("A1:B1").Style.Color = Color.DarkGray
            ' Set the font color of cells A1 to B1
            sheet.Range("A1:B1").Style.Font.Color = Color.White
            ' Set the vertical alignment of cells A1 to B1
            sheet.Range("A1:B1").Style.VerticalAlignment = VerticalAlignType.Center
            ' Set the horizontal alignment of cells A1 to B1
            sheet.Range("A1:B1").Style.HorizontalAlignment = HorizontalAlignType.Center
            'Apply a number format for cells  B2:C5
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
