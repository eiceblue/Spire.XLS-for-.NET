Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace WriteComment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WriteComment.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a font object for regular comment
            Dim font As ExcelFont = workbook.CreateFont()
            font.FontName = "Arial"
            font.Size = 11
            font.KnownColor = ExcelColors.Orange

            ' Create a font object for blue text
            Dim fontBlue As ExcelFont = workbook.CreateFont()
            fontBlue.KnownColor = ExcelColors.LightBlue

            ' Create a font object for green text
            Dim fontGreen As ExcelFont = workbook.CreateFont()
            fontGreen.KnownColor = ExcelColors.LightGreen

            ' Select a specific cell range (B11)
            Dim range As CellRange = sheet.Range("B11")

            ' Set the text content of the cell
            range.Text = "Regular comment"

            ' Set the text content of the comment associated with the cell
            range.Comment.Text = "Regular comment"

            ' Autofit the width of the column containing the cell range
            range.AutoFitColumns()

            ' Select a different cell range (B12)
            range = sheet.Range("B12")

            ' Set the text content of the cell
            range.Text = "Rich text comment"

            ' Apply the font to the first 17 characters of the cell's rich text
            range.RichText.SetFont(0, 16, font)

            ' Autofit the width of the column containing the cell range
            range.AutoFitColumns()

            ' Set the text content of the comment associated with the cell
            range.Comment.RichText.Text = "Rich text comment"

            ' Apply the green font to the first 5 characters of the comment's rich text
            range.Comment.RichText.SetFont(0, 4, fontGreen)

            ' Apply the blue font to characters 6 to 10 of the comment's rich text
            range.Comment.RichText.SetFont(5, 9, fontBlue)

            ' Specify the output file name
            Dim result As String = "WriteComment_result.xlsx"

            ' Save the workbook to the specified output file in Excel 2007 format
            workbook.SaveToFile(result, ExcelVersion.Version2007)
            ' Release the resources used by the workbook
            workbook.Dispose()
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
