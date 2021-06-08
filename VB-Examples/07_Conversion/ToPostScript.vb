Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ToPostScript
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'create a workbook
			Dim workbook As New Workbook()

			'load an excel document
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ToPostScript.xlsx")

			Dim result As String = "Result.ps"
			'convert to ODS file
			workbook.SaveToFile(result, FileFormat.PostScript)

			'view the document
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
