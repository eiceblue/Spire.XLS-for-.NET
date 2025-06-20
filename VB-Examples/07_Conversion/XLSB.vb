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
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing XLSB file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\XLSB.xlsb")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Export the contents of the worksheet to a DataTable and set it as the data source for a data grid
            Me.dataGrid1.DataSource = sheet.ExportDataTable()
			Me.btnSave.Enabled = True
		End Sub

		Private Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click

            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Insert the data from the data grid's data source (assumed to be a DataTable) into the worksheet starting at cell B2
            sheet.InsertDataTable(CType(Me.dataGrid1.DataSource, DataTable), True, 1, 1, -1, -1)

            ' Define cell styles for alternating rows
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

            ' Apply alternating row styles based on the row index
            For Each range As CellRange In sheet.AllocatedRange.Rows
                If range.Row Mod 2 = 0 Then
                    range.CellStyleName = evenStyle.Name
                Else
                    range.CellStyleName = oddStyle.Name
                End If
            Next range

            ' Set cell style for the header row
            Dim styleHeader As CellStyle = sheet.AllocatedRange.Rows(0).Style
            styleHeader.Borders(BordersLineType.EdgeLeft).LineStyle = LineStyleType.Thin
            styleHeader.Borders(BordersLineType.EdgeRight).LineStyle = LineStyleType.Thin
            styleHeader.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
            styleHeader.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Thin
            styleHeader.VerticalAlignment = VerticalAlignType.Center
            styleHeader.KnownColor = ExcelColors.Green
            styleHeader.Font.KnownColor = ExcelColors.White
            styleHeader.Font.IsBold = True

            ' Auto-fit the columns and rows in the worksheet
            sheet.AllocatedRange.AutoFitColumns()
            sheet.AllocatedRange.AutoFitRows()

            ' Set the height of the first row to 20 points
            sheet.Rows(0).RowHeight = 20

            ' Save the workbook to a XLSB file with Excel 2010 format
            workbook.SaveToFile("sample.xlsb", ExcelVersion.Xlsb2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

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