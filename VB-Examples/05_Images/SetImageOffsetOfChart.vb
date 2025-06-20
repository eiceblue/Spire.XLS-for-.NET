Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace SetImageOffsetOfChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load an Excel document from the specified file.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Add a new worksheet with the name "Contrast".
            Dim sheet1 As Worksheet = workbook.Worksheets.Add("Contrast")

            'Add a clustered column chart to the "Contrast" worksheet.
            Dim chart1 As Chart = sheet1.Charts.Add(ExcelChartType.ColumnClustered)

            'Set the data range for the chart to cells D1 to E8 in the "sheet" worksheet.
            chart1.DataRange = sheet.Range("D1:E8")

            'Specify that the series data for the chart is not automatically generated from the range.
            chart1.SeriesDataFromRange = False

            'Sets the starting column index of the chart.
            chart1.LeftColumn = 1
            'Sets the starting row index of the chart.
            chart1.TopRow = 11
            'Sets the ending column index of the chart.
            chart1.RightColumn = 8
            'Sets the ending row index of the chart.
            chart1.BottomRow = 33

            'Set the custom picture fill for the chart area using the specified image file and transparency setting.
            chart1.ChartArea.Fill.CustomPicture(Image.FromFile("Background.png"), "None")

            'Disable tiling of the picture within the chart area.
            chart1.ChartArea.Fill.Tile = False

            'Set the picture stretch values for left, top, right, and bottom sides of the chart area.
            chart1.ChartArea.Fill.PicStretch.Left = 20
            chart1.ChartArea.Fill.PicStretch.Top = 20
            chart1.ChartArea.Fill.PicStretch.Right = 5
            chart1.ChartArea.Fill.PicStretch.Bottom = 5

            'Specify the filename to save the modified workbook.
            Dim result As String = "Result-SetImageOffsetOfChart.xlsx"

            'Save the workbook to the specified file using Excel 2013 format.
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
            ExcelDocViewer(result)
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
