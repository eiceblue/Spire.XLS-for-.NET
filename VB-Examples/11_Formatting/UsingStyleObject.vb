Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace UsingStyleObject
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Add a new worksheet with the name "new sheet"
            Dim sheet As Worksheet = workbook.Worksheets.Add("new sheet")

            ' Get the cell range B1
            Dim cell As CellRange = sheet.Range("B1")

            ' Set the text of the cell to "Hello Spire!"
            cell.Text = "Hello Spire!"

            ' Create a new cell style named "newStyle" and add it to the workbook's styles
            Dim style As CellStyle = workbook.Styles.Add("newStyle")

            ' Set the vertical alignment and horizontal alignment of the style to center
            style.VerticalAlignment = VerticalAlignType.Center
            style.HorizontalAlignment = HorizontalAlignType.Center

            ' Set the font color of the style to blue
            style.Font.Color = Color.Blue

            ' Enable shrinking the text to fit within the cell
            style.ShrinkToFit = True

            ' Set the bottom border color and line style of the style
            style.Borders(BordersLineType.EdgeBottom).Color = Color.GreenYellow
            style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Medium

            ' Apply the style to the cell
            cell.Style = style

            ' Apply the style to cell B4 and set its text to "Test"
            sheet.Range("B4").Style = style
            sheet.Range("B4").Text = "Test"

            ' Apply the style to cell C3 using the style's name property and set its text
            sheet.Range("C3").CellStyleName = style.Name
            sheet.Range("C3").Text = "Welcome to use Spire.XLS"

            ' Apply the style to cell D4
            sheet.Range("D4").Style = style

            ' Save the workbook to a file named "UsingStyleObject_result.xlsx" in Excel 2010 format
            Dim result As String = "UsingStyleObject_result.xlsx"
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
