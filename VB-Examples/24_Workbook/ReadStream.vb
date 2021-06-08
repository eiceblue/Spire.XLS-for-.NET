Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace ReadStream

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()

			'Open excel from a stream
			Dim fileStream As FileStream = File.OpenRead("..\..\..\..\..\..\Data\ReadStream.xlsx")
			fileStream.Seek(0, SeekOrigin.Begin)

			workbook.LoadFromStream(fileStream)

			workbook.SaveToFile("ReadStream_result.xlsx",ExcelVersion.Version2013)
			ExcelDocViewer("ReadStream_result.xlsx")
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
