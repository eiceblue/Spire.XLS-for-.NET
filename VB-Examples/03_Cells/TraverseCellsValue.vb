Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace TraverseCellsValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CellValues.xlsx")

			'Get first worksheet of the workbook
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Get the cell range collection 
			Dim cellRangeCollection() As CellRange = worksheet.Cells

			'Create StringBuilder to save 
			Dim content As New StringBuilder()
			content.AppendLine("Values of the first sheet:")

			'Traverse cells value
			For Each cellRange As CellRange In cellRangeCollection
				'Set string format for displaying
				Dim result As String = String.Format("Cell: " & cellRange.RangeAddress & "   Value: " & cellRange.Value)

				'Add result string to StringBuilder
				content.AppendLine(result)
			Next cellRange
			'String for output file 
			Dim outputFile As String = "Output.txt"

			'Save them to a txt file
			File.WriteAllText(outputFile, content.ToString())

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
