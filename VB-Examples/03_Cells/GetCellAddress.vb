Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace GetCellAddress
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim builder As New StringBuilder()

			'Get a cell range
			Dim range As CellRange = sheet.Range("A1:B5")

			'Get address of range
			Dim address As String = range.RangeAddressLocal
			builder.AppendLine("Address of range: " & address)

			'Get the cell count of range
			Dim count As Integer = range.CellsCount
			builder.AppendLine("Cell count of range: " & count.ToString())

			'Get the address of the entire column of range
			Dim entireColAddress As String = range.EntireColumn.RangeAddressLocal
			builder.AppendLine("Address of entire column of the range: " & entireColAddress)

			'Get the address of the entire row of range
			Dim entireRowAddress As String = range.EntireRow.RangeAddressLocal
			builder.AppendLine("Address of entire row of the range " & entireRowAddress)

			'Write to txt file
			Dim output As String = "GetCellAddress_out.txt"
			File.WriteAllText(output, builder.ToString())

			'Launch the txt file
			ExcelDocViewer(output)
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
