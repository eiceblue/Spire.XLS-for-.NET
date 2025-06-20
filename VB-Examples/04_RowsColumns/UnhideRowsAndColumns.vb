Imports Spire.Xls

Namespace UnhideRowsAndColumns

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new instance of the Workbook class.
            Dim workbook As New Workbook()
            'Load the Excel document from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\HideRowsAndColumns.xls")
            'Retrieve the first worksheet from the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            'Hide the second column of the worksheet.
            worksheet.HideColumn(2)
            'Hide the fourth row of the worksheet.
            worksheet.HideRow(4)
            ' Specify the file name
            Dim output As String = "HideRowsAndColumns.xlsx"
            'Save the modified workbook to a file with the specified name and Excel version.
            workbook.SaveToFile(output, ExcelVersion.Version2010)
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
