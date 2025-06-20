Imports Spire.Xls

Namespace AddTrendline

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "ChartSample2.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample2.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart (index 0) on the worksheet
            Dim chart As Chart = sheet.Charts(0)

            ' Set the chart title to "Logarithmic Trendline"
            chart.ChartTitle = "Logarithmic Trendline"

            ' Add a logarithmic trendline to the first series of the chart
            chart.Series(0).TrendLines.Add(TrendLineType.Logarithmic)

            ' Get the second chart (index 1) on the worksheet
            Dim chart1 As Chart = sheet.Charts(1)

            ' Set the chart title to "Moving Average Trendline"
            chart1.ChartTitle = "Moving Average Trendline"

            ' Add a moving average trendline to the first series of the chart
            chart1.Series(0).TrendLines.Add(TrendLineType.Moving_Average)

            ' Get the third chart (index 2) on the worksheet
            Dim chart2 As Chart = sheet.Charts(2)

            ' Set the chart title to "Linear Trendline"
            chart2.ChartTitle = "Linear Trendline"

            ' Add a linear trendline to the first series of the chart
            chart2.Series(0).TrendLines.Add(TrendLineType.Linear)

            ' Get the fourth chart (index 3) on the worksheet
            Dim chart3 As Chart = sheet.Charts(3)

            ' Set the chart title to "Exponential Trendline"
            chart3.ChartTitle = "Exponential Trendline"

            ' Add an exponential trendline to the first series of the chart
            chart3.Series(0).TrendLines.Add(TrendLineType.Exponential)

            ' Specify the output filename for the modified workbook
            Dim output As String = "AddTrendline.xlsx"

            ' Save the modified workbook to a new Excel file with the Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
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
