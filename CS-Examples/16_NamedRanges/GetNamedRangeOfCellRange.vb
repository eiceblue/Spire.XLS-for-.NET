Imports Spire.Xls
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.IO
Imports System.Reflection.Emit

Namespace GetNamedRangeOfCellRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			Dim outputFile As String = "GetNamedRangeOfCellRange.txt"

			' Create a new workbook object
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AllNamedRanges.xlsx")

			' Determine whether NamedRange exists in Range A7:D7
			Dim result = workbook.Worksheets(0).Range("A7:D7").GetNamedRange()
			File.WriteAllText(outputFile, "A7:D7---" & result.Name & vbCrLf)

			' Determine whether NamedRange exists in Range A4:D4
			Dim result1 = workbook.Worksheets(0).Range("A4:D4").GetNamedRange()
			File.AppendAllText(outputFile, "A4:D4---" & result1.Name & vbCrLf)

			' Determine whether NamedRange exists in cell C14
			Dim result2 = workbook.Worksheets(0).Range("C14").GetNamedRange()
			If result2 Is Nothing Then
				File.AppendAllText(outputFile, "C14 cell does not have NameRange")
			End If

			workbook.CalculateAllValue()

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			FileViewer(outputFile)

			Me.Close()
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

		Private Sub label1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles label1.Click

		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
