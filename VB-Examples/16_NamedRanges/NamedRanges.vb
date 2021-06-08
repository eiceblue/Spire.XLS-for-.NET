Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace NamedRanges
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\NamedRanges.xlsx")
			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Creating a named range
			Dim NamedRange As INamedRange = workbook.NameRanges.Add("NewNamedRange")
			'Setting the range of the named range
			NamedRange.RefersToRange = sheet.Range("A8:E12")

			Dim result As String = "NamedRanges_result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2013)
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
