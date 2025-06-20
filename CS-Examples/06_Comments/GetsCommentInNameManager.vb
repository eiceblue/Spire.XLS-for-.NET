Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet

Namespace GetsCommentInNameManager
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new Workbook object
			Dim workbook As New Workbook()

			' Load the Excel file "GetNotesInformation.xlsx" from a specific path
			workbook.LoadFromFile("..\..\..\..\..\..\Data\GetNotesInformation.xlsx")

			' Access the NameRanges property of the workbook
			Dim nameManager As INameRanges = workbook.NameRanges

			' Create a StringBuilder to store the result
			Dim stringBuilder As New StringBuilder()

			' Iterate through each name in the NameRanges collection
			For i As Integer = 0 To nameManager.Count - 1
				' Get the XlsName object at index i
				Dim name As XlsName = CType(nameManager(i), XlsName)

				' Append the name and comment value to the StringBuilder
				stringBuilder.Append("Name: " & name.Name & ", Comment: " & name.CommentValue & vbCrLf)
			Next i

			' Write the result to a text file named "GetsCommentInNameManager_result.txt"
			File.WriteAllText("GetsCommentInNameManager_result.txt", stringBuilder.ToString())

			' Dispose of the workbook object
			workbook.Dispose()
		        ExcelDocViewer("GetsCommentInNameManager_result.txt")
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
