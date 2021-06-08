Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace GetPageCount

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

			Dim pageInfoList = workbook.GetSplitPageInfo()
			Dim sb As New StringBuilder()
			For i As Integer = 0 To workbook.Worksheets.Count - 1
				Dim sheetname As String = workbook.Worksheets(i).Name
				Dim pagecount As Integer = pageInfoList(i).Count
				sb.AppendLine(sheetname & "'s page count is: " & pagecount)
			Next i

			'Save the document
			Dim output As String = "GetPageCount.txt"
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
