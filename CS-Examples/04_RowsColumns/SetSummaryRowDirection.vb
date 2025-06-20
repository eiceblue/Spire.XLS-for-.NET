Imports Spire.Xls

Namespace SetSummaryRowDirection

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new Workbook object.
            Dim workbook As New Workbook()

            'Loads the Excel document from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample2.xlsx")

            'Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Group columns 1 to 4 in the worksheet, with expanded state set to True
            sheet.GroupByColumns(1, 4, True)

            ' Set the option to display summary columns on the right side of the detail columns.
            sheet.PageSetup.IsSummaryColumnRight = True

            'Specify the name for the resulting file.
            Dim output As String = "SetSummaryColumnDirection.xlsx"
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
