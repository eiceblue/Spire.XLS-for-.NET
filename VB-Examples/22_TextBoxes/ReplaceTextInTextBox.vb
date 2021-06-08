Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace ReplaceTextInTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ReplaceTextInTextBox.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim tag As String = "TAG_1$TAG_2"
			Dim replace As String = "Spire.XLS for .NET$Spire.XLS for JAVA"

			For i As Integer = 0 To tag.Split("$"c).Length - 1
				'Replace text in textbox
				ReplaceTextInTextBox(sheet, "<" & tag.Split("$"c)(i) & ">", replace.Split("$"c)(i))
			Next i

			'Save the document
			Dim output As String = "ReplaceTextInTextBox_out.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(output)
		End Sub
		Private Sub ReplaceTextInTextBox(ByVal sheet As Worksheet, ByVal sFind As String, ByVal sReplace As String)
			For i As Integer = 0 To sheet.TextBoxes.Count - 1
				Dim tb As ITextBox = sheet.TextBoxes(i)
				If Not String.IsNullOrEmpty(tb.Text) Then
					If tb.Text.Contains(sFind) Then
						tb.Text = tb.Text.Replace(sFind, sReplace)
					End If
				End If
			Next i
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
