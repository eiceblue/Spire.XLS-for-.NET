Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace DetectProtection

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load the document from disk
			Dim input As String = "..\..\..\..\..\..\Data\ProtectedWorkbook.xlsx"
			'Detect if the Excel workbook is password protected
			Dim value As Boolean = Workbook.IsPasswordProtected(input)

			If value Then
				textBox1.Text = "Yes"
			Else
				textBox1.Text = "No"
			End If

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
