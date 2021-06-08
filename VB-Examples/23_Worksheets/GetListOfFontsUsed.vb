Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports System.IO

Namespace GetListOfFontsUsed
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load a excel document
			workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz.xlsx")

			Dim fonts As New List(Of ExcelFont)()

			'Loop all sheets of workbook
			For Each sheet As Worksheet In workbook.Worksheets
				For r As Integer = 0 To sheet.Rows.Length - 1
					For c As Integer = 0 To sheet.Rows(r).CellList.Count - 1
						'Get the font of cell and add it to list
						fonts.Add(sheet.Rows(r).CellList(c).Style.Font)
					Next c
				Next r
			Next sheet
			Dim strB As New StringBuilder()

			For Each font As ExcelFont In fonts
				strB.AppendLine(String.Format("FontName:{0}; FontSize:{1}",font.FontName,font.Size))
			Next font

			Dim result As String = "GetListOfFontsUsed_result.txt"

			File.WriteAllText(result, strB.ToString())
			'View the document
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
	End Class
End Namespace
