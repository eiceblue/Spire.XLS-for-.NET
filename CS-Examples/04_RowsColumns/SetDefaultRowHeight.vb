Imports Spire.Xls

Namespace SetDefaultRowHeight

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new Workbook object.
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the Workbook object.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CommonTemplate.xlsx")

            'Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Set the default row height to 30.
            sheet.DefaultRowHeight = 30

            'Specify the name for the resulting file.
            Dim output As String = "SetDefaultRowHeight.xlsx"
            'Save the workbook to a file with the specified name and Excel version (in this case, Excel 2013).
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
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
