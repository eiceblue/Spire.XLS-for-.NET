Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetHeightAndWidth
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SetHeightAndWidth.xls")

			Dim worksheet As Worksheet = workbook.Worksheets(0)
			' Setting the width to 30
			worksheet.SetColumnWidth(4, 30)
			' Setting the height to 30
			worksheet.SetRowHeight(4,30)

			Dim result As String="SetHeightAndWidth_out.xlsx"
			workbook.SaveToFile(result,ExcelVersion.Version2010)
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
