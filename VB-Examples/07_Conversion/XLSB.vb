Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace XLSB
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnLoad_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLoad.Click
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\XLSB.xlsb")
			'Initailize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Me.dataGrid1.DataSource = sheet.ExportDataTable()
			Me.btnSave.Enabled = True
		End Sub

		Private Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click

			Dim workbook As New Workbook()

			'Initailize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.InsertDataTable(CType(Me.dataGrid1.DataSource, DataTable), True,1, 1, -1, -1)

			'Sets body style
			Dim oddStyle As CellStyle = workbook.Styles.Add("oddStyle")
			oddStyle.Borders(BordersLineType.EdgeLeft).LineStyle = LineStyleType.Thin
			oddStyle.Borders(BordersLineType.EdgeRight).LineStyle = LineStyleType.Thin
			oddStyle.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
			oddStyle.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Thin
			oddStyle.KnownColor = ExcelColors.LightGreen1

			Dim evenStyle As CellStyle = workbook.Styles.Add("evenStyle")
			evenStyle.Borders(BordersLineType.EdgeLeft).LineStyle = LineStyleType.Thin
			evenStyle.Borders(BordersLineType.EdgeRight).LineStyle = LineStyleType.Thin
			evenStyle.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
			evenStyle.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Thin
			evenStyle.KnownColor = ExcelColors.LightTurquoise

			For Each range As CellRange In sheet.AllocatedRange.Rows
				If range.Row Mod 2 = 0 Then
					range.CellStyleName = evenStyle.Name
				Else
					range.CellStyleName = oddStyle.Name
				End If
			Next range

			'Sets header style
			Dim styleHeader As CellStyle = sheet.AllocatedRange.Rows(0).Style
			styleHeader.Borders(BordersLineType.EdgeLeft).LineStyle = LineStyleType.Thin
			styleHeader.Borders(BordersLineType.EdgeRight).LineStyle = LineStyleType.Thin
			styleHeader.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
			styleHeader.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Thin
			styleHeader.VerticalAlignment = VerticalAlignType.Center
			styleHeader.KnownColor = ExcelColors.Green
			styleHeader.Font.KnownColor = ExcelColors.White
			styleHeader.Font.IsBold = True

			sheet.AllocatedRange.AutoFitColumns()
			sheet.AllocatedRange.AutoFitRows()

			sheet.Rows(0).RowHeight = 20

			workbook.SaveToFile("sample.xlsb", ExcelVersion.Xlsb2010)
			ExcelDocViewer("sample.xlsb")
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