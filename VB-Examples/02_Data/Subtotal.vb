Imports Spire.Xls

Namespace Subtotal
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()
            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Subtotal.xlsx")
            'Gets the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Defines a range of cells from A1 to B18.
            Dim range As CellRange = sheet.Range("A1:B18")

            'Applies subtotals to the selected data range with the following settings:
            '- Group by column index 0 (column A).
            '- Summarize values in column index 1 (column B) using the Sum function.
            '- Show the resulting subtotals.
            '- Replace the current subtotals.
            '- Not add page break between groups
            '- Add summary of data
            sheet.Subtotal(range, 0, New Integer() {1}, SubtotalTypes.Sum, True, False, True)

            'Specifies the filename for the resulting Excel file.
            Dim result As String = "Subtotal_Out.xlsx"
            'Saves the modified workbook to a file with the specified filename and Excel version.
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
