Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace FindCellsWithStyleName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load the document from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the cell style name
			Dim styleName As String = sheet.Range("A1").CellStyleName

			Dim ranges As CellRange = sheet.AllocatedRange
			For Each cc As CellRange In ranges
				'Find the cells which have the same style name
				If cc.CellStyleName = styleName Then
					'Set value
					cc.Value = "Same style"
				End If
			Next cc

			Dim result As String = "FindCellsWithStyleName_result.xlsx"
			'Save and launch result file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
			FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
