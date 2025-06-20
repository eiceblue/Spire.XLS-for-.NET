Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CreateMultiLevelChart

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

            ' Write data to cells
            sheet.Range("A1").Text = "Main Category"
            sheet.Range("A2").Text = "Fruit"
            sheet.Range("A6").Text = "Veggies"
            sheet.Range("B1").Text = "Sub Category"
            sheet.Range("B2").Text = "Bananas"
            sheet.Range("B3").Text = "Oranges"
            sheet.Range("B4").Text = "Pears"
            sheet.Range("B5").Text = "Grapes"
            sheet.Range("B6").Text = "Carrots"
            sheet.Range("B7").Text = "Potatoes"
            sheet.Range("B8").Text = "Celery"
            sheet.Range("B9").Text = "Onions"
            sheet.Range("C1").Text = "Value"
            sheet.Range("C2").Value = "52"
            sheet.Range("C3").Value = "65"
            sheet.Range("C4").Value = "50"
            sheet.Range("C5").Value = "45"
            sheet.Range("C6").Value = "64"
            sheet.Range("C7").Value = "62"
            sheet.Range("C8").Value = "89"
            sheet.Range("C9").Value = "57"

            ' Merge cells in column A for Fruit category
            sheet.Range("A2:A5").Merge()

            ' Merge cells in column A for Veggies category
            sheet.Range("A6:A9").Merge()

            ' Autofit columns 1 (A) and 2 (B) to fit the content
            sheet.AutoFitColumn(1)
            sheet.AutoFitColumn(2)

            ' Add a clustered bar chart to the worksheet
            Dim chart As Chart = sheet.Charts.Add(ExcelChartType.BarClustered)

            ' Set the chart title to "Value"
            chart.ChartTitle = "Value"

            ' Remove fill from the plot area of the chart
            chart.PlotArea.Fill.FillType = ShapeFillType.NoFill

            ' Delete the legend in the chart
            chart.Legend.Delete()

            ' Set the position of the chart within the worksheet using column and row indices
            chart.LeftColumn = 5
            chart.TopRow = 1
            chart.RightColumn = 14

            ' Set the data range for the chart
            chart.DataRange = sheet.Range("C2:C9")

            ' Specify that series data will not be derived from the range
            chart.SeriesDataFromRange = False

            ' Get the first series in the chart
            Dim serie As ChartSerie = chart.Series(0)

            ' Set the category labels for the series using a range that includes both columns A and B
            serie.CategoryLabels = sheet.Range("A2:B9")

            ' Enable multi-level category axis labels
            chart.PrimaryCategoryAxis.MultiLevelLable = True

            ' Specify the output file name as "CreateMultiLevelChart.xlsx"
            Dim output As String = "CreateMultiLevelChart.xlsx"

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
