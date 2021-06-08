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
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)
			Dim sheet1 As Worksheet = workbook.Worksheets.Add("Contrast")

			'Add chart1 and background image to sheet1 as comparision.
			Dim chart1 As Chart = sheet1.Charts.Add(ExcelChartType.ColumnClustered)
			chart1.DataRange = sheet.Range("D1:E8")
			chart1.SeriesDataFromRange = False

			'Chart Position.
			chart1.LeftColumn = 1
			chart1.TopRow = 11
			chart1.RightColumn = 8
			chart1.BottomRow = 33

			'Add picture as background.
			chart1.ChartArea.Fill.CustomPicture(Image.FromFile("..\..\..\..\..\..\Data\Background.png"), "None")
			
	  	chart1.ChartArea.Fill.Tile = False
	  	
			'Set the image offset.		
			chart1.ChartArea.Fill.PicStretch.Left = 20
			chart1.ChartArea.Fill.PicStretch.Top = 20
			chart1.ChartArea.Fill.PicStretch.Right = 5
			chart1.ChartArea.Fill.PicStretch.Bottom = 5

			Dim result As String = "Result-SetImageOffsetOfChart.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

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
