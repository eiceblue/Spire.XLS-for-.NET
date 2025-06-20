Imports Spire.Xls

Namespace SetCellFillPattern

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CommonTemplate.xlsx")
            'Retrieves the first worksheet from the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            'Sets the background color of the cells in the range "B7:F7" to yellow.
            worksheet.Range("B7:F7").Style.Color = Color.Yellow

            'Sets the fill pattern of the cells in the range "B8:F8" to a 125% gray pattern.
            worksheet.Range("B8:F8").Style.FillPattern = ExcelPatternType.Percent125Gray

            'Specifies the name of the output file.
            Dim output As String = "SetCellFillPattern.xlsx"
            'Saves the workbook to the specified file in Excel 2013 format.
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
