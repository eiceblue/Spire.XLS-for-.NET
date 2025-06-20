Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace FormatARow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a new cell style named "newStyle"
            Dim style As CellStyle = workbook.Styles.Add("newStyle")

            ' Set the vertical alignment of the style to center
            style.VerticalAlignment = VerticalAlignType.Center

            ' Set the horizontal alignment of the style to center
            style.HorizontalAlignment = HorizontalAlignType.Center

            ' Set the font color of the style to blue
            style.Font.Color = Color.Blue

            ' Enable the "shrink to fit" option for the style
            style.ShrinkToFit = True

            ' Set the bottom border color of the style to orange-red
            style.Borders(BordersLineType.EdgeBottom).Color = Color.OrangeRed

            ' Set the line style of the bottom border of the style to dotted
            style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Dotted

            ' Apply the "newStyle" cell style to all cells in row 1
            sheet.Rows(1).CellStyleName = style.Name

            ' Set the text of all cells in row 1 to "Test"
            sheet.Rows(1).Text = "Test"

            ' Specify the file name for saving the workbook
            Dim result As String = "FormatARow_result.xlsx"

            ' Save the modified workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)
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
