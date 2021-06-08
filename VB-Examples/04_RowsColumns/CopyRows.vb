Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace CopyRows
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load an excel file including pivot table
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Copying.xls")

			Dim sheet1 As Worksheet = workbook.Worksheets(1)
			Dim sheet2 As Worksheet = workbook.Worksheets(0)

			'Copy the first row to the third row in the same sheet
			sheet1.Copy(sheet1.Rows(0), sheet1.Rows(2), True, True, True)

			'Copy the first row to the second row in the different sheet
			sheet1.Copy(sheet1.Rows(0), sheet2.Rows(1), True, True, True)

			Dim result As String = "CopyRows_result.xlsx"

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
