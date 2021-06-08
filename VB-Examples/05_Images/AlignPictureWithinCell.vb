Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace AlignPictureWithinCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.Range("A1").Text = "Align Picture Within A Cell:"
			sheet.Range("A1").Style.VerticalAlignment = VerticalAlignType.Top

			'Insert an image to the specific cell.
			Dim picPath As String = "..\..\..\..\..\..\Data\SpireXls.png"
			Dim picture As ExcelPicture = sheet.Pictures.Add(1, 1, picPath)

			'Adjust the column width and row height so that the cell can contain the picture.
			sheet.Columns(0).ColumnWidth = 40
			sheet.Rows(0).RowHeight = 200

			'Vertically and horizontally align the image.
			picture.LeftColumnOffset = 100
			picture.TopRowOffset = 25

			Dim result As String = "Result-AlignPictureWithinCell.xlsx"

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
