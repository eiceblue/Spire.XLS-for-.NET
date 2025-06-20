Imports System.IO
Imports System.Text
Imports Spire.Xls
Imports Spire.Xls.Core

Namespace ExtractTrendline

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Declare a new workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from a file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample4.xlsx")

            ' Get the first chart from the first worksheet in the workbook
            Dim chart As Chart = workbook.Worksheets(0).Charts(0)

            ' Get the trendline object from the second series of the chart
            Dim trendLine As IChartTrendLine = chart.Series(1).TrendLines(0)

            ' Get the formula of the trendline
            Dim formula As String = trendLine.Formula

            ' Create a StringBuilder to store the output text
            Dim sb As New StringBuilder()

            ' Append the equation information to the StringBuilder
            sb.AppendLine("The equation is: " & formula)

            ' Specify the output file name
            Dim output As String = "ExtractTrendline.txt"

            ' Write the contents of the StringBuilder to the output file
            File.WriteAllText(output, sb.ToString())
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
