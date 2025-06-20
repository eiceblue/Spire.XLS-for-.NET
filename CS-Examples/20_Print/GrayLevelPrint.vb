Imports Spire.Xls

Namespace GrayLevelPrint
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_3.xlsx")

			' Set the GrayLevelForPrint to true
			workbook.ConverterSetting.GrayLevelForPrint = True

			' Print this document
			workbook.PrintDocument.Print()

			' Dispose of the workbook object to release resources
			workbook.Dispose()

		End Sub
		Private Sub OutputViewer(ByVal fileName As String)
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
