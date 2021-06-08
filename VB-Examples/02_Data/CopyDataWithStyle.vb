Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopyDataWithStyle
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the default first worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Set the values for some cells.
			Dim cells As CellRange = worksheet.Range("A1:J50")
			For i As Integer = 1 To 10
				For j As Integer = 1 To 8
					Dim text As String = String.Format((i - 1).ToString() & "," & (j - 1).ToString())
					cells(i, j).Text = text
				Next j
			Next i
			'Get a source range (A1:D3).
			Dim srcRange As CellRange = worksheet.Range("A1:D3")

			'Create a style object.
			Dim style As CellStyle = workbook.Styles.Add("style")

			'Specify the font attribute.
			style.Font.FontName = "Calibri"

			'Specify the shading color.
			style.Font.Color = Color.Red

			'Specify the border attributes.
			style.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
			style.Borders(BordersLineType.EdgeTop).Color = Color.Blue
			style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Thin
			style.Borders(BordersLineType.EdgeBottom).Color = Color.Blue
			style.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
			style.Borders(BordersLineType.EdgeTop).Color = Color.Blue
			style.Borders(BordersLineType.EdgeRight).LineStyle = LineStyleType.Thin
			style.Borders(BordersLineType.EdgeRight).Color = Color.Blue
			srcRange.CellStyleName = style.Name

			'Set the destination range
			Dim destRange As CellRange = worksheet.Range("A12:D14")

			'Copy the range data with style
			srcRange.Copy(destRange, True, True)

			'String for output file 
			Dim outputFile As String = "Output.xlsx"

			'Save the file
			workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

			'Launching the output file.
			Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
