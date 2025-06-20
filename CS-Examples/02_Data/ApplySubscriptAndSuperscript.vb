Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ApplySubscriptAndSuperscript
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Excel workbook object.
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)
            ' Set the text of cell B2 as "This is an example of Subscript:".
            sheet.Range("B2").Text = "This is an example of Subscript:"
            ' Set the text of cell D2 as "This is an example of Superscript:".
            sheet.Range("D2").Text = "This is an example of Superscript:"

            ' Get the range of cell B3.
            Dim range As CellRange = sheet.Range("B3")
            ' Set the rich text value of the cell as "R100-0.06".
            range.RichText.Text = "R100-0.06"

            ' Create a new font object.
            Dim font As ExcelFont = workbook.CreateFont()
            ' Set the IsSubscript property of the font to true (for subscript).
            font.IsSubscript = True
            ' Set the color of the font to green.
            font.Color = Color.Green

            ' Apply the specified font to the range of characters from index 4 to 8.
            range.RichText.SetFont(4, 8, font)

            ' Get the range of cell D3.
            range = sheet.Range("D3")
            ' Set the rich text value of the cell as "a2 + b2 = c2".
            range.RichText.Text = "a2 + b2 = c2"

            ' Create another font object.
            font = workbook.CreateFont()
            ' Set the IsSuperscript property of the font to true (for superscript).
            font.IsSuperscript = True

            ' Apply the specified font to the range of characters at index 1.
            range.RichText.SetFont(1, 1, font)
            ' Apply the specified font to the range of characters at index 6.
            range.RichText.SetFont(6, 6, font)
            ' Apply the specified font to the range of characters at index 11.
            range.RichText.SetFont(11, 11, font)
            ' Auto-fit the column widths to fit the content.
            sheet.AllocatedRange.AutoFitColumns()
            ' Specify the output file name.
            Dim result As String = "Result-ApplySubscriptAndSuperscript.xlsx"

            ' Save the modified workbook to the specified file with Excel 2013 format.
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
