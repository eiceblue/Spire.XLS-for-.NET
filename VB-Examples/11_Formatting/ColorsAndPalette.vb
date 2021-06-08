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
			'Create a workbook
			Dim workbook As New Workbook()

			'Adding Orchid color to the palette at 60th index
			workbook.ChangePaletteColor(Color.Orchid, 60)

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim cell As CellRange = sheet.Range("B2")
			cell.Text = "Welcome to use Spire.XLS"

			'Set the Orchid (custom) color to the font
			cell.Style.Font.Color = Color.Orchid
			cell.Style.Font.Size = 20
			cell.AutoFitColumns()
			cell.AutoFitRows()

			Dim result As String = "ColorsAndPalette_result.xlsx"

			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

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
