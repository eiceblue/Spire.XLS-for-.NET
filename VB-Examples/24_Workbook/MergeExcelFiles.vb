Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace MergeExcelFiles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim files As New List(Of String)()
			files.Add("..\..\..\..\..\..\Data\MergeExcelFiles-1.xlsx")
			files.Add("..\..\..\..\..\..\Data\MergeExcelFiles-2.xls")
			files.Add("..\..\..\..\..\..\Data\MergeExcelFiles-3.xlsx")

			Dim newbook As New Workbook()
			newbook.Version = ExcelVersion.Version2013
			'Clear all worksheets
			newbook.Worksheets.Clear()

			'Create a workbook
			Dim tempbook As New Workbook()

			For Each file As String In files
				'Load the file
				tempbook.LoadFromFile(file)
				For Each sheet As Worksheet In tempbook.Worksheets
					'Copy every sheet in a workbook
					newbook.Worksheets.AddCopy(sheet,WorksheetCopyType.CopyAll)
				Next sheet
			Next file

			'Save the file
			newbook.SaveToFile("MergeExcelFiles.xlsx", ExcelVersion.Version2010)
			ExcelDocViewer("MergeExcelFiles.xlsx")
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
