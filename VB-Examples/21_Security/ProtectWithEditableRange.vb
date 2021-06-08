Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ProtectWithEditableRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook and load a file
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ProtectWithEditableRange.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Define the specified ranges to allow users to edit while sheet is protected
			sheet.AddAllowEditRange("EditableRanges", sheet.Range("B4:E12"))

			'Protect worksheet with a password.
			sheet.Protect("TestPassword", SheetProtectionType.All)

			Dim result As String = "ProtectWithEditableRange_result.xlsx"
			'Save the document and launch it
			workbook.SaveToFile(result, ExcelVersion.Version2010)
			ExcelDocViewer(result)
		End Sub

		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
		Private Sub btnAbout_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAbout.Click
			Close()
		End Sub
	End Class
End Namespace
