Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace SetFontForLegendAndDataTable

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load a workbook file named "ChartSample1.xlsx" from a specific file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim ws As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = ws.Charts(0)

            ' Create a new font object
            Dim font As ExcelFont = workbook.CreateFont()

            ' Set the font size to 14 and color to Red
            font.Size = 14.0
            font.Color = Color.Red

            ' Set the font for the legend text area of the chart
            chart.Legend.TextArea.SetFont(font)

            ' Iterate through each series in the chart
            For Each cs As ChartSerie In chart.Series
                ' Set the font for data labels of each data point in the series
                cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font)
            Next cs

            ' Define the output file name as "SetFontForLegendAndDataTable.xlsx"
            Dim output As String = "SetFontForLegendAndDataTable.xlsx"

            ' Save the modified workbook to the specified output file with Excel 2013 format
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
