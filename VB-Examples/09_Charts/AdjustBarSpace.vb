Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace AdjustBarSpace

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "ChartSample1.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim ws As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet
            Dim chart As Chart = ws.Charts(0)

            ' Iterate through each series in the chart
            For Each cs As ChartSerie In chart.Series
                ' Set the gap width of the series to 200
                cs.Format.Options.GapWidth = 200
                ' Set the overlap of the series to 0
                cs.Format.Options.Overlap = 0
            Next cs

            ' Specify the output file name as "AdjustBarSpace.xlsx"
            Dim output As String = "AdjustBarSpace.xlsx"

            ' Save the modified workbook to the specified file path, using the Excel 2013 format
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
