Imports Spire.Xls

Namespace SetFontForTitleAndAxis

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
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart in the worksheet
            Dim chart As Chart = worksheet.Charts(0)

            ' Set the color, size, and font name for the chart title area
            chart.ChartTitleArea.Color = Color.Blue
            chart.ChartTitleArea.Size = 20.0
            chart.ChartTitleArea.FontName = "Arial"

            ' Set the color and size for the primary value axis font
            chart.PrimaryValueAxis.Font.Color = Color.Gold
            chart.PrimaryValueAxis.Font.Size = 10.0

            ' Set the font name, color, and size for the primary category axis font
            chart.PrimaryCategoryAxis.Font.FontName = "Arial"
            chart.PrimaryCategoryAxis.Font.Color = Color.Red
            chart.PrimaryCategoryAxis.Font.Size = 20.0

            ' Specify the output file name
            Dim output As String = "SetFontForTitleAndAxis.xlsx"

            ' Save the modified workbook to a new file
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
