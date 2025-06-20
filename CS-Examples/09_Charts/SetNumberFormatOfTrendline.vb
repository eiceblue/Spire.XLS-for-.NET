Imports System.IO
Imports System.Text
Imports Spire.Xls
Imports Spire.Xls.Core

Namespace SetNumberFormatOfTrendline

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample4.xlsx")

            ' Get the chart object from the first worksheet of the workbook
            Dim chart As Chart = workbook.Worksheets(0).Charts(0)

            ' Get the trendline object from the first series of the chart
            Dim trendLine As IChartTrendLine = chart.Series(1).TrendLines(0)

            ' Set the number format of the data label on the trendline to "#,##0.00"
            trendLine.DataLabel.NumberFormat = "#,##0.00"

            ' Specify the output file name for saving the modified workbook
            Dim output As String = "SetNumberFormatOfTrendline_out.xlsx"

            ' Save the workbook to the specified output file path in Excel 2013 format
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
