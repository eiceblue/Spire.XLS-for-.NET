Imports Spire.Xls

Namespace RemovePictureBorder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads an Excel document from the specified file.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PictureBorder.xlsx")

            'Gets the first worksheet from the workbook.
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            'Retrieves the first picture from the first worksheet.
            Dim picture As ExcelPicture = sheet1.Pictures(0)

            'Remove the picture border
            'Method-1:
            'Sets the visibility of the border line for the picture to false.
            picture.Line.Visible = False

            'Method-2:
            'Sets the weight (thickness) of the border line for the picture to 0.
            'picture.Line.Weight = 0;

            'Specifies the filename to save the modified workbook.
            Dim result As String = "RemovePictureBorder.xlsx"

            'Saves the workbook to the specified file using Excel 2010 format.
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
