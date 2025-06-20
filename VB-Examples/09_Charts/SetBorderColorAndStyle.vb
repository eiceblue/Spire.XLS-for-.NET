Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Charts

Namespace SetBorderColorAndStyle

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample3.xlsx")

            ' Get the first worksheet in the workbook
            Dim ws As Worksheet = workbook.Worksheets(0)

            ' Get the first chart in the worksheet
            Dim chart As Chart = ws.Charts(0)

            ' Set the line weight of the first data point in the first series of the chart
            TryCast(chart.Series(0).DataPoints(0).DataFormat.LineProperties, XlsChartBorder).CustomLineWeight = 2.5F

            ' Set the color of the line of the first data point in the first series of the chart
            TryCast(chart.Series(0).DataPoints(0).DataFormat.LineProperties, XlsChartBorder).Color = Color.Red

            ' Save the modified workbook to a new file
            Dim output As String = "SetBorderColorAndStyle.xlsx"
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
