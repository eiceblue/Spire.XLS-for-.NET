Imports Spire.Xls

Namespace DeleteLegend

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "ChartSample1.xlsx"
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSample1.xlsx")

            ' Get the first worksheet from the loaded workbook and assign it to a variable named "sheet"
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart from the worksheet and assign it to a variable named "chart"
            Dim chart As Chart = sheet.Charts(0)

            ' Delete the legend from the chart (commented out in the code)
            ' chart.Legend.Delete();

            ' Delete specific legend entries from the chart
            chart.Legend.LegendEntries(0).Delete()
            chart.Legend.LegendEntries(1).Delete()

            ' Specify the output file name
            Dim output As String = "DeleteLegend.xlsx"

            ' Save the workbook to a file named "DeleteLegend.xlsx" in Excel 2013 format
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
