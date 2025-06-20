Imports Spire.Xls

Namespace ToPdfWithCustomPageSize
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

            ' Iterate through each worksheet in the workbook
            For Each sheet As Worksheet In workbook.Worksheets

                ' Set a custom paper size for the current worksheet
                sheet.PageSetup.SetCustomPaperSize(100.0F, 100.0F)
            Next sheet

            ' Specify the filename for the resulting PDF file
            Dim result As String = "result.pdf"

            ' Save the workbook as a PDF file with the specified filename and format
            workbook.SaveToFile(result, FileFormat.PDF)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer(result)
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
