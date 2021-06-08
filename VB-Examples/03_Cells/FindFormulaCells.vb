Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace FindFormulaCells
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\FindCellsSample.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Find the cells that contain formula "=SUM(A11,A12)"
			Dim ranges() As CellRange = sheet.FindAll("=SUM(A11,A12)", FindType.Formula, ExcelFindOptions.None)

			'Create a string builder
			Dim builder As New StringBuilder()

			'Append the address of found cells to builder
			If ranges.Length <> 0 Then
				For Each range As CellRange In ranges
					Dim address As String = range.RangeAddress
					builder.AppendLine("The address of found cell is: " & address)
				Next range
			Else
				builder.AppendLine("No cell contain the formula")
			End If

			'Save to txt file
			Dim result As String = "FindFormulaCells_out.txt"
			File.WriteAllText(result, builder.ToString())

			'Launch the file
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
