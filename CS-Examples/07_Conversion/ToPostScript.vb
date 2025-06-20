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
            ' Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ToPostScript.xlsx")

            ' Specify the file name
            Dim result As String = "result.ps"

            ' Save the workbook as a PostScript file.
            workbook.SaveToFile(result, FileFormat.PostScript)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
