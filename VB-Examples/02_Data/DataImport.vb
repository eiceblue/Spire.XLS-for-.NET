Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace DataImport
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Initialize a Worksheet object by getting the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Insert data from a DataTable into the worksheet, starting from cell (1,1)
            sheet.InsertDataTable(CType(Me.dataGrid1.DataSource, DataTable), True, 1, 1, -1, -1)

            ' Set the body style for alternate rows
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

            ' Apply alternating row styles based on row number
            For Each range As CellRange In sheet.AllocatedRange.Rows
                If range.Row Mod 2 = 0 Then
                    range.CellStyleName = evenStyle.Name
                Else
                    range.CellStyleName = oddStyle.Name
                End If
            Next range

            ' Set the header style for the first row
            Dim styleHeader As CellStyle = sheet.Rows(0).Style
            styleHeader.Borders(BordersLineType.EdgeLeft).LineStyle = LineStyleType.Thin
            styleHeader.Borders(BordersLineType.EdgeRight).LineStyle = LineStyleType.Thin
            styleHeader.Borders(BordersLineType.EdgeTop).LineStyle = LineStyleType.Thin
            styleHeader.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Thin
            styleHeader.VerticalAlignment = VerticalAlignType.Center
            styleHeader.KnownColor = ExcelColors.Green
            styleHeader.Font.KnownColor = ExcelColors.White
            styleHeader.Font.IsBold = True

            ' Auto-fit columns and rows within the allocated range
            sheet.AllocatedRange.AutoFitColumns()
            sheet.AllocatedRange.AutoFitRows()

            ' Set the row height for the first row
            sheet.Rows(0).RowHeight = 20

            ' Specify the file name for saving the workbook
            Dim result As String = "DataImport_out.xls"

            ' Save the modified workbook to the specified file path
            workbook.SaveToFile(result)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer(result)
		End Sub
		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\DataImport.xls")
			'Initailize worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Me.dataGrid1.DataSource = sheet.ExportDataTable()
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
