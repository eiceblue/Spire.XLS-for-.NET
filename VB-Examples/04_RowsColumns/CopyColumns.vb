Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace CopyColumns
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

			Dim sheet1 As Worksheet = workbook.Worksheets(0)
			Dim sheet2 As Worksheet = workbook.Worksheets(1)

			'Copy the first column to the third column in the same sheet
			sheet1.Copy(sheet1.Columns(0),sheet1.Columns(2),True,True,True)

			'Copy the first column to the second column in the different sheet
			sheet1.Copy(sheet1.Columns(0), sheet2.Columns(1), True, True, True)

			Dim result As String = "CopyColumns_result.xlsx"

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
