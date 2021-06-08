Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace RegisterAddInFunction
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim input As String = "..\..\..\..\..\..\Data\Test.xlam"

			'Create a workbook
			Dim workbook As New Workbook()
			'Register AddIn function
			workbook.AddInFunctions.Add(input, "TEST_UDF")
			workbook.AddInFunctions.Add(input, "TEST_UDF1")
			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Call AddIn function
			sheet.Range("A1").Formula = "=TEST_UDF()"
			sheet.Range("A2").Formula = "=TEST_UDF1()"

			Dim result As String = "RegisterAddInFunction_result.xlsx"

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
