Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace FormatCellsWithStyle
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load the document from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

			'Create a style
			Dim style As CellStyle = workbook.Styles.Add("newStyle")
			'Set the shading color
			style.Color = Color.DarkGray
			'Set the font color
			style.Font.Color = Color.White
			'Set font name
			style.Font.FontName = "Times New Roman"
			'Set font size
			style.Font.Size = 12
			'Set bold for the font
			style.Font.IsBold = True
			'Set text rotation
			style.Rotation = 45
			'Set alignment
			style.HorizontalAlignment = HorizontalAlignType.Center
			style.VerticalAlignment = VerticalAlignType.Center

			'Set the style for the specific range
			workbook.Worksheets(0).Range("A1:J1").CellStyleName = style.Name

			'Save and launch result file
			Dim result As String = "result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
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
