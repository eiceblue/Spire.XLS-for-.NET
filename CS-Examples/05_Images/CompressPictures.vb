Imports Spire.Xls

Namespace CompressPictures
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Load an Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CompressPictures.xlsx")

            'Assign the first worksheet to the "sheet1" variable.
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            'Iterate through each worksheet in the workbook.
            For Each sheet As Worksheet In workbook.Worksheets
                'Iterate through each picture in the worksheet.
                For Each picture As ExcelPicture In sheet.Pictures
                    'Compresse the picture with a compression level of 50%.
                    picture.Compress(50)
                Next picture
            Next sheet

            'Specifie the file name for the resulting workbook.
            Dim result As String = "CompressPictures_result.xlsx"
            'Save the workbook to a file with the specified name and version.
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
