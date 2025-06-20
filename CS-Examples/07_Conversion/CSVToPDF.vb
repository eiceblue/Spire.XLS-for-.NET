Imports Spire.Xls

Namespace CSVToPDF

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the CSV file from the specified path with the specified delimiter (","),
            ' starting at row 1 and column 1
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CSVSample.csv", ",", 1, 1)

            ' Set the SheetFitToPage property of the ConverterSetting to True
            workbook.ConverterSetting.SheetFitToPage = True

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Auto-fit columns in the worksheet
            For i As Integer = 1 To sheet.Columns.Length - 1
                sheet.AutoFitColumn(i)
            Next i

            ' Specify the output file name for saving as PDF
            Dim output As String = "CSVToPDF.pdf"

            ' Save the workbook to a PDF file
            workbook.SaveToFile(output, FileFormat.PDF)
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
