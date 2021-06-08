Imports Spire.Xls
Imports System.IO

Namespace GetFreezePaneRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load an excel file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\GetFreezePaneRange.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)
			Dim rowIndex As Integer
			Dim colIndex As Integer

			'The row and column index of the frozen pane is passed through the out parameter. 
			'If it returns to 0, it means that it is not frozen
			sheet.GetFreezePanes(rowIndex, colIndex)

			Dim range As String = "Row index: " & rowIndex & ", column index: " & colIndex

			'Save the document and launch it
			Dim result As String = "GetFreezePaneCellRange_result.txt"
			File.WriteAllText(result, range)
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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
