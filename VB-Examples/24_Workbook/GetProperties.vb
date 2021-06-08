Imports System.IO
Imports System.Text
Imports Spire.Xls
Imports Spire.Xls.Collections
Imports Spire.Xls.Core

Namespace GetProperties

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

			'Get the general excel properties
			Dim properties1 As BuiltInDocumentProperties = workbook.DocumentProperties
			Dim sb As New StringBuilder()
			sb.AppendLine("Excel Properties:")
			For i As Integer = 0 To properties1.Count - 1
				Dim name As String = properties1(i).Name
				Dim value As String = properties1(i).Value.ToString()
				sb.AppendLine(name & ": " & value)
			Next i
			sb.AppendLine()

			'Get the custom properties
			Dim properties2 As ICustomDocumentProperties = workbook.CustomDocumentProperties
			sb.AppendLine("Custom Properties:")
			For i As Integer = 0 To properties2.Count - 1
				Dim name As String = properties2(i).Name
				Dim value As String = properties2(i).Value.ToString()
				sb.AppendLine(name & ": " & value)
			Next i
			'Save the document
			Dim output As String = "GetProperties.txt"
			File.WriteAllText(output, sb.ToString())

			'Launch the Excel file
			ExcelDocViewer(output)
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
