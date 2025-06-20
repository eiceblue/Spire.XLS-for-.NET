Imports Spire.Xls

Namespace LoadAndSaveFileWithMacro

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\MacroSample.xls")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text of cell A5 to "This is a simple test!"
            sheet.Range("A5").Text = "This is a simple test!"

            ' Specify the output file name
            Dim output As String = "LoadAndSaveFileWithMacro.xls"

            ' Save the workbook to the specified file path with Excel version compatibility set to Version97to2003
            workbook.SaveToFile(output, ExcelVersion.Version97to2003)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
            ExcelDocViewer(output)
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
