Imports Spire.Xls

Namespace DeleteMultipleRowsAndColumns

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads the Excel document from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CommonTemplate1.xlsx")

            'Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Deletes four rows starting from the fifth row.
            sheet.DeleteRow(5, 4)

            'Deletes two columns starting from the second column.
            sheet.DeleteColumn(2, 2)

            'Specifies the name of the resulting Excel file.
            Dim output As String = "DeleteMultipleRowsAndColumns.xlsx"
            'Saves the modified workbook to a file with the specified name and Excel version.
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
