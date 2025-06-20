Imports Spire.Xls

Namespace SetSummaryColumnDirection

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to work with Excel files
            Dim workbook As New Workbook()

            ' Load the Excel file "WorksheetSample1.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

            ' Get the first worksheet from the Workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Group rows 1 to 4 in the worksheet, with expanded state set to True
            sheet.GroupByRows(1, 4, True)

            ' Set the option to display summary rows below the detail rows
            sheet.PageSetup.IsSummaryRowBelow = False

            ' Specify the output file name as "SetSummaryRowDirection.xlsx"
            Dim output As String = "SetSummaryRowDirection.xlsx"
            ' Save the modified workbook to the specified output file in the Excel 2013 format
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
