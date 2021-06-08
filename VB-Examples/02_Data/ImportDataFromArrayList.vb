Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace ImportDataFromArrayList
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Create an empty worksheet
			workbook.CreateEmptySheets(1)

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create an ArrayList object
			Dim list As New ArrayList()

			'Add strings in list
			list.Add("Spire.Doc for .NET")
			list.Add("Spire.XLS for .NET")
			list.Add("Spire.PDF for .NET")
			list.Add("Spire.Presentation for .NET")

			'Insert array list in worksheet 
			sheet.InsertArrayList(list, 1, 1, True)

			'Save the Excel file
			Dim result As String = "ImportDataFromArrayList_out.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the Excel file
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
