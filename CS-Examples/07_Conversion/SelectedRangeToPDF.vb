Imports Spire.Xls

Namespace SelectedRangeToPDF

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

            ' Add a new worksheet to the workbook
            workbook.Worksheets.Add("newsheet")

            ' Copy the range A9:E15 from the first worksheet to the second worksheet
            workbook.Worksheets(0).Range("A9:E15").Copy(workbook.Worksheets(1).Range("A9:E15"), False, True)

            ' Auto-fit columns in the range A9:E15 of the second worksheet
            workbook.Worksheets(1).Range("A9:E15").AutoFitColumns()

            ' Specify the output file name for saving the second worksheet as PDF
            Dim output As String = "SelectedRangeToPDF.pdf"

            ' Save the second worksheet to a PDF file
            workbook.Worksheets(1).SaveToPdf(output)
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
