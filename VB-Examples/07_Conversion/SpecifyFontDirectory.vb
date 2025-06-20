Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SpecifyFontDirectory
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new instance of Workbook within a "Using" block for automatic disposal
			Using workbook As New Workbook()
				' Load the Excel file from the specified path
				workbook.LoadFromFile("..\..\..\..\..\..\Data\ToPDFSample.xlsx")

				' Specify the directory containing custom font files to be used in the workbook
				workbook.CustomFontFileDirectory = New String() {("..\..\..\..\..\..\Data\Font")}

				' Save the workbook to a PDF file with the specified output file name
				workbook.SaveToFile("result.pdf", FileFormat.PDF)
			End Using

			ExcelDocViewer("result.pdf")
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
