Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace ReadHyperlinks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadHyperlinks.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the address of the first hyperlink in the worksheet
            textBox1.Text = sheet.HyperLinks(0).Address

            ' Get the address of the second hyperlink in the worksheet
            textBox2.Text = sheet.HyperLinks(1).Address
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
