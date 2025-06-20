Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Charts

Namespace SetLegendBackgroundColor

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim ws As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = ws.Charts(0)

            ' Cast the Legend's FrameFormat to XlsChartFrameFormat
            Dim x As XlsChartFrameFormat = TryCast(chart.Legend.FrameFormat, XlsChartFrameFormat)

            ' Set the fill type of the legend frame to solid color
            x.Fill.FillType = ShapeFillType.SolidColor

            ' Set the foreground color (background color) of the legend frame to SkyBlue
            x.ForeGroundColor = Color.SkyBlue

            ' Specify the output file name for saving the modified workbook
            Dim output As String = "SetLegendBackgroundColor.xlsx"

            ' Save the workbook to the specified output file with Excel 2013 format
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
