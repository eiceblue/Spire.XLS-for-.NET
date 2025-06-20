Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace SetCommentFillColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Gets the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Create a new Excel font.
            Dim font As ExcelFont = workbook.CreateFont()
            font.FontName = "Arial"
            font.Size = 11
            font.KnownColor = ExcelColors.Orange

            'Get the CellRange object for cell A1.
            Dim range As CellRange = sheet.Range("A1")

            'Set the comment text for the cell.
            range.Comment.Text = "This is a comment"

            'Set the font for the comment's rich text.
            range.Comment.RichText.SetFont(0, (range.Comment.Text.Length - 1), font)

            'Set the fill type of the comment's shape fill to solid color.
            range.Comment.Fill.FillType = ShapeFillType.SolidColor

            'Set the foreground color (fill color) for the comment's shape.
            range.Comment.Fill.ForeColor = Color.SkyBlue

            'Set the visibility of the comment to true.
            range.Comment.Visible = True

            'Specify the filename to save the modified workbook.
            Dim result As String = "SetCommentFillColor_out.xlsx"

            'Save the workbook to the specified file using Excel 2013 format.
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
