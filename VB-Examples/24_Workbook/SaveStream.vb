Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports Spire.Xls

Namespace SaveStream
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SaveStream.xls")

			'Save an excel workbook to stream
			Dim fileStream As New FileStream("SaveStream.xlsx", FileMode.Create)
			workbook.SaveToStream(fileStream,FileFormat.Version2010)

			fileStream.Close()

			ExcelDocViewer("SaveStream.xlsx")
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
