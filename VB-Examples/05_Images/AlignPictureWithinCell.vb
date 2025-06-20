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
            'Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Assign the first worksheet to the "sheet" variable.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Set the text content of cell A1 to "Align Picture Within A Cell:".
            sheet.Range("A1").Text = "Align Picture Within A Cell:"
            'Sets the vertical alignment of the cell A1 to the top.
            sheet.Range("A1").Style.VerticalAlignment = VerticalAlignType.Top

            'Specifie the file path of the image.
            Dim picPath As String = "..\..\..\..\..\..\Data\SpireXls.png"
            'Adds an image to the specified cell (1, 1).
            Dim picture As ExcelPicture = sheet.Pictures.Add(1, 1, picPath)

            'Set the width of the first column to 40.
            sheet.Columns(0).ColumnWidth = 40
            'Sets the height of the first row to 200.
            sheet.Rows(0).RowHeight = 200

            'Set the horizontal offset of the image.
            picture.LeftColumnOffset = 100
            'Sets the vertical offset of the image.
            picture.TopRowOffset = 25

            'Specify the file name for the resulting workbook.
            Dim result As String = "Result-AlignPictureWithinCell.xlsx"

            'Save the workbook to a file with the specified name and version.
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
