Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace MakeCellActive
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			'Read an Excel file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz.xlsx")

			'Get the 2nd sheet
			Dim sheet As Worksheet = workbook.Worksheets(1)

			'Set the 2nd sheet as an active sheet.
			sheet.Activate()

			'Set B2 cell as an active cell in the worksheet.
			sheet.SetActiveCell(sheet.Range("B2"))

			'Set the B column as the first visible column in the worksheet.
			sheet.FirstVisibleColumn = 1

			'Set the 2nd row as the first visible row in the worksheet.
			sheet.FirstVisibleRow = 1

			Dim result As String = "MakeCellActive_result.xlsx"

			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

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
