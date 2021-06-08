Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports System.IO

Namespace GetExcelPaperDimensions
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

			Dim content As New StringBuilder()

			'Get the dimensions of A2 paper.
			sheet.PageSetup.PaperSize = PaperSizeType.A2Paper
			content.AppendLine("A2Paper: " & sheet.PageSetup.PageWidth & " x " & sheet.PageSetup.PageHeight)

			'Get the dimensions of A3 paper.
			sheet.PageSetup.PaperSize = PaperSizeType.PaperA3
			content.AppendLine("PaperA3: " & sheet.PageSetup.PageWidth & " x " & sheet.PageSetup.PageHeight)

			'Get the dimensions of A4 paper.
			sheet.PageSetup.PaperSize = PaperSizeType.PaperA4
			content.AppendLine("PaperA4: " & sheet.PageSetup.PageWidth & " x " & sheet.PageSetup.PageHeight)

			'Get the dimensions of paper letter.
			sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter
			content.AppendLine("PaperLetter: " & sheet.PageSetup.PageWidth & " x " & sheet.PageSetup.PageHeight)

			Dim result As String = "Result-GetExcelPaperDimensions.txt"

			'Save to file.
			File.WriteAllText(result, content.ToString())

			'Launch the file.
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
