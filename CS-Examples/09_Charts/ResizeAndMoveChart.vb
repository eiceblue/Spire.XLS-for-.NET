Imports Spire.Xls

Namespace ResizeAndMoveChart

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

            ' Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet.
            Dim chart As Chart = sheet.Charts(0)

            ' Set the column index where the chart will be positioned.
            chart.LeftColumn = 5

            ' Set the row index where the chart will be positioned.
            chart.TopRow = 1

            ' Set the width of the chart in pixels.
            chart.Width = 500

            ' Set the height of the chart in pixels.
            chart.Height = 350

            ' Specify the output file name for the modified workbook.
            Dim output As String = "ResizeAndMoveChart.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2013 format.
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
