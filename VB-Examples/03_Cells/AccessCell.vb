Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace AccessCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AccessCell.xlsx")

			Dim builder As New StringBuilder()

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Access cell by its name
			Dim range1 As CellRange = sheet.Range("A1")
			builder.AppendLine("Value of range1: " & range1.Text)

			'Access cell by index of row and column
			Dim range2 As CellRange = sheet.Range(2,1)
			builder.AppendLine("Value of range2: " & range2.Text)

			'Access cell in cell collection
			Dim range3 As CellRange = sheet.Cells(2)
			builder.AppendLine("Value of range3: " & range3.Text)

			'Save to txt file
			Dim result As String="AccessCell_out.txt"
			File.WriteAllText(result, builder.ToString())

			'Launch the txt file
			OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal fileName As String)
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
