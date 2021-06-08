Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace UseExplicitLineBreaks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the first default worksheet
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			'Specify a cell range
			Dim c5 As CellRange = sheet1.Range("C5")

			'Set the cell width for specified range
			sheet1.SetColumnWidth(c5.Column, 70)

			'Put the string value with explicit line breaks
			c5.Value = "Spire.XLS for .NET is a professional Excel .NET API" & vbLf & " that can be used to create, read, " & vbLf & "write, convert and print Excel files in any type " & vbLf & "of .NET(C#, VB.NET, ASP.NET, .NET Core) application. " & vbLf & "Spire.XLS for .NET offers object model" & vbLf & " Excel API for speeding up Excel programming in .NET platform -" & vbLf & " create new Excel documents from template, edit existing " & vbLf & "Excel documents and " & vbLf & "convert Excel files."

			'Set Text wrap
			c5.IsWrapText = True

			'String for output file 
			Dim outputFile As String = "Output.xlsx"

			'Save the file
			workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

			'Launching the output file.
			Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
