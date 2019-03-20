Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetBorder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SetBorder.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the cell range where you want to apply border style
			Dim cr As CellRange = sheet.Range(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn)

			'Apply border style 
			cr.Borders.LineStyle = LineStyleType.Double
			cr.Borders(BordersLineType.DiagonalDown).LineStyle = LineStyleType.None
			cr.Borders(BordersLineType.DiagonalUp).LineStyle = LineStyleType.None
			cr.Borders.Color = Color.CadetBlue

			'Save the document
			workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010)

			'Launch the Excel file
			ExcelDocViewer(workbook.FileName)
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
