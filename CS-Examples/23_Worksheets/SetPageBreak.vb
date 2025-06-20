Imports Spire.Xls

Namespace SetPageBreak

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a horizontal page break at cell A8 on the worksheet
            sheet.HPageBreaks.Add(sheet.Range("A8"))

            ' Add another horizontal page break at cell A14 on the worksheet
            sheet.HPageBreaks.Add(sheet.Range("A14"))

            ' Add vertical page breaks
            ' sheet.VPageBreaks.Add(sheet.Range("B1"))
            ' sheet.VPageBreaks.Add(sheet.Range("C1"))

            ' Set the view mode of the first worksheet to Page Break Preview
            workbook.Worksheets(0).ViewMode = ViewMode.Preview

            ' Specify the output filename for the modified workbook
            Dim output As String = "SetPageBreak.xlsx"

            ' Save the workbook to a file in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

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
