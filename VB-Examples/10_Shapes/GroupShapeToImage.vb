Imports Spire.Xls
Imports System.Drawing.Imaging

Namespace GroupShapeToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook class
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\GroupShapeToImage.xlsx")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Create an instance of SaveShapeTypeOption class to specify saving options
            Dim saveShapeTypeOption As New SaveShapeTypeOption()

            ' Set the option to save group shapes
            saveShapeTypeOption.SaveGroupShape = True

            ' Save the group shapes in the worksheet as images and get a list of Bitmap objects
            Dim images As List(Of Bitmap) = worksheet.SaveShapesToImage(saveShapeTypeOption)

            ' Iterate through the list of images
            For i As Integer = 0 To images.Count - 1

                ' Generate a unique image file name
                Dim imageFile As String = String.Format("Image-{0}.png", i)

                ' Save the image as PNG file
                images(i).Save(imageFile, ImageFormat.Png)

            Next i

            ' Dispose the workbook object to release resources
            workbook.Dispose()
        End Sub
		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
