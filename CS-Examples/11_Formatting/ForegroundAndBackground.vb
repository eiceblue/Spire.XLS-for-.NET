Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet

Namespace ForegroundAndBackground
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a new cell style named "newStyle1"
            Dim style As CellStyle = workbook.Styles.Add("newStyle1")

            ' Set the fill pattern of the interior to vertical stripes
            style.Interior.FillPattern = ExcelPatternType.VerticalStripe

            ' Set the background color of the gradient to green
            style.Interior.Gradient.BackKnownColor = ExcelColors.Green

            ' Set the foreground color of the gradient to yellow
            style.Interior.Gradient.ForeKnownColor = ExcelColors.Yellow

            ' Apply the "newStyle1" cell style to cell B2
            sheet.Range("B2").CellStyleName = style.Name

            ' Set the text of cell B2 to "Test"
            sheet.Range("B2").Text = "Test"

            ' Set the row height of cell B2 to 30
            sheet.Range("B2").RowHeight = 30

            ' Set the column width of cell B2 to 50
            sheet.Range("B2").ColumnWidth = 50

            ' Add a new cell style named "newStyle2"
            style = workbook.Styles.Add("newStyle2")

            ' Set the fill pattern of the interior to thin horizontal stripes
            style.Interior.FillPattern = ExcelPatternType.ThinHorizontalStripe

            ' Set the foreground color of the gradient to red
            style.Interior.Gradient.ForeKnownColor = ExcelColors.Red

            ' Apply the "newStyle2" cell style to cell B4
            sheet.Range("B4").CellStyleName = style.Name

            ' Set the row height of cell B4 to 30
            sheet.Range("B4").RowHeight = 30

            ' Set the column width of cell B4 to 60
            sheet.Range("B4").ColumnWidth = 60

            ' Save the modified workbook to a new file named "ForegroundAndBackground_result.xlsx" using Excel 2010 format
            Dim result As String = "ForegroundAndBackground_result.xlsx"
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
