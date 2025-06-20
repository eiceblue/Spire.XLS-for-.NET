Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.Collections

Namespace ColorsAndPalette
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Change the palette color to Orchid at index 60
            workbook.ChangePaletteColor(Color.Orchid, 60)

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the cell range B2
            Dim cell As CellRange = sheet.Range("B2")

            ' Set the text in the cell
            cell.Text = "Welcome to use Spire.XLS"

            ' Set the font color, size, and autofit the columns and rows of the cell
            cell.Style.Font.Color = Color.Orchid
            cell.Style.Font.Size = 20
            cell.AutoFitColumns()
            cell.AutoFitRows()

            ' Save the workbook to a file
            Dim result As String = "ColorsAndPalette_result.xlsx"
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
