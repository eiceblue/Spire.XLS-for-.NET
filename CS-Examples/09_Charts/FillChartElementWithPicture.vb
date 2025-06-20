Imports Spire.Xls

Namespace FillChartElementWithPicture

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim ws As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = ws.Charts(0)

            ' Set a custom picture as the fill for the chart area, specifying the image file path and fill style
            chart.ChartArea.Fill.CustomPicture(Image.FromFile("..\..\..\..\..\..\Data\background.png"), "None")

            ' Set the transparency of the plot area fill to 0.9 (90% transparent)
            chart.PlotArea.Fill.Transparency = 0.9

            ' Save the modified workbook to a new file named "FillChartElementWithPicture.xlsx"
            Dim output As String = "FillChartElementWithPicture.xlsx"
            workbook.SaveToFile(output, ExcelVersion.Version2010)
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
