Imports Spire.Xls

Namespace DetectVBAMacros

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\MacroSample.xls")

			'Detect if the Excel file contains VBA macros
			Dim hasMacros As Boolean = False
			hasMacros = workbook.HasMacros
			If hasMacros Then
				Me.textBox1.Text = "Yes"

			Else
				Me.textBox1.Text = "No"
			End If
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
