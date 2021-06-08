Imports Spire.Xls
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Namespace ChartToEMFImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a Workbook from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToEMFImage.xlsx")

			'Save chart as Emf image
			Using stream As New MemoryStream()
				workbook.SaveChartAsEmfImage(workbook.Worksheets(0), 0, stream)
				File.WriteAllBytes("EmfImage.emf", stream.ToArray())
			End Using

			'Launch the file
			ExcelDocViewer("EmfImage.emf")
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
