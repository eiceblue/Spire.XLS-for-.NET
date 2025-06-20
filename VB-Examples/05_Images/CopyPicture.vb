Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopyPicture
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            'Create a new worksheet named "DestSheet" in the workbook.
            Dim destinationSheet As Worksheet = workbook.Worksheets.Add("DestSheet")

            'Retrieve the first picture from the first worksheet.
            Dim sourcePicture As ExcelPicture = sheet1.Pictures(0)

            'Extract the image data from the source picture.
            Dim image As Image = sourcePicture.Picture

            'Insert the image at cell coordinates (2, 2) in the destination worksheet.
            destinationSheet.Pictures.Add(2, 2, image)

            'Specify the name of the output file.
            Dim outputFile As String = "Output.xlsx"

            'Save the modified workbook to the specified output file using Excel 2013 format.
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
