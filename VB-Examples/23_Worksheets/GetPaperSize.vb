Imports System.IO
Imports System.Text
Imports Spire.Xls

Namespace GetPaperSize

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample2.xlsx")

			Dim sb As New StringBuilder()
			For Each sheet As Worksheet In workbook.Worksheets
				Dim width As Double = sheet.PageSetup.PageWidth
				Dim height As Double = sheet.PageSetup.PageHeight
				sb.AppendLine(sheet.Name)
				sb.AppendLine("Width: " & width & vbTab & "Height: " & height)
				sb.AppendLine()
			Next sheet

			'Save to Text file
			Dim output As String = "GetPaperSize.txt"
			File.WriteAllText(output, sb.ToString())

			'Launch the file
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
