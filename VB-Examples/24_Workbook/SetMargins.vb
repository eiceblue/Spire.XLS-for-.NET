Imports Spire.Xls

Namespace SetMargins

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from a file using a relative path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

            ' Access the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the top, bottom, left, and right margins of the worksheet's page setup
            sheet.PageSetup.TopMargin = 0.3
            sheet.PageSetup.BottomMargin = 1
            sheet.PageSetup.LeftMargin = 0.2
            sheet.PageSetup.RightMargin = 1

            ' Set the header and footer margins of the worksheet's page setup
            sheet.PageSetup.HeaderMarginInch = 0.1
            sheet.PageSetup.FooterMarginInch = 0.5

            ' Specify the output file name for saving the modified workbook
            Dim output As String = "SetMargins.xlsx"

            ' Save the workbook to a file with the specified output name and Excel version
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
