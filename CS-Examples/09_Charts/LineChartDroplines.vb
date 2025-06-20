Imports Spire.Xls

Namespace LineChartDroplines
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            Using workbook As New Workbook()
                ' Load the Excel file from disk.
                workbook.LoadFromFile("..\..\..\..\..\..\Data\LineChartDroplines.xlsx")

                ' Get the first worksheet in the workbook.
                Dim worksheet As Worksheet = workbook.Worksheets(0)

                ' Get the first chart in the worksheet.
                Dim chart As Chart = worksheet.Charts(0)

                ' Enable drop lines for the first series of the chart.
                chart.Series(0).HasDroplines = True

                ' Save the modified document to "result.xlsx" using the Excel 2013 format.
                workbook.SaveToFile("result.xlsx", FileFormat.Version2013)
            End Using
            ExcelDocViewer("result.xlsx")
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
