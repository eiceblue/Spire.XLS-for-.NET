Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.Text
Imports System.IO

Namespace GetAllNamedRange
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim sb As New StringBuilder()

			'Load the document from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AllNamedRanges.xlsx")

			'Get all named range
			Dim ranges As INameRanges = workbook.NameRanges
			For Each nameRange As INamedRange In ranges
				sb.Append(nameRange.Name & vbCrLf)
			Next nameRange

			'Save and launch result file
			Dim result As String = "result.txt"
			File.WriteAllText(result, sb.ToString())
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
