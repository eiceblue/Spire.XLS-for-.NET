Imports Spire.Xls

Namespace SetChartBackgroundColor

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

            ' Get the first worksheet in the workbook
            Dim ws As Worksheet = workbook.Worksheets(0)

            ' Get the first chart in the worksheet
            Dim chart As Chart = ws.Charts(0)

            ' Set the foreground color of the chart area to LightYellow
            chart.ChartArea.ForeGroundColor = Color.LightYellow

            ' Save the modified workbook to a new file
            Dim output As String = "SetChartBackgroundColor.xlsx"
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
