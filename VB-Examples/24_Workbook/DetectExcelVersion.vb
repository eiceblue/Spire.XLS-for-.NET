Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace DetectExcelVersion
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Files
			Dim files() As String = { "..\..\..\..\..\..\Data\ExcelSample97_N.xls", "..\..\..\..\..\..\Data\ExcelSample_N1.xlsx", "..\..\..\..\..\..\Data\ExcelSample_N.xlsb" }

			Dim builder As New StringBuilder()

			For Each file As String In files
				'Create a workbook
				Dim workbook As New Workbook()

				'Load the document
				workbook.LoadFromFile(file)

				'Get the version
				Dim version As ExcelVersion = workbook.Version

				builder.AppendLine(version.ToString())
			Next file

			'Save to txt file
			Dim result As String = "DetectExcelVersion_out.txt"
			File.WriteAllText(result, builder.ToString())

			'Launch the file
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
