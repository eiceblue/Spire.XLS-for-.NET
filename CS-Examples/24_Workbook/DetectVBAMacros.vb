Imports Spire.Xls

Namespace DetectVBAMacros

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\MacroSample.xls")

            ' Declare and initialize a variable to store whether the workbook has macros
            Dim hasMacros As Boolean = False

            ' Check if the workbook has macros
            hasMacros = workbook.HasMacros

            ' If the workbook has macros, set the text of textBox1 to "Yes"
            If hasMacros Then
                Me.textBox1.Text = "Yes"
            Else
                ' If the workbook does not have macros, set the text of textBox1 to "No"
                Me.textBox1.Text = "No"
            End If

            ' Release the resources used by the workbook
            workbook.Dispose()
        End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

	End Class
End Namespace
