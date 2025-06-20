Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace UOSToExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\input.uos", ExcelVersion.UOS)

            ' Save the workbook to an xlsx file format.
            workbook.SaveToFile("output.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()
            FileViewer("output.xlsx")
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
