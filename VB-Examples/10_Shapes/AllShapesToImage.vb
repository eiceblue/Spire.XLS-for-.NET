Imports Spire.Xls
Imports System.Drawing.Imaging

Namespace AllShapesToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Shape.xlsx")

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Create an instance of SaveShapeTypeOption to specify saving options
            Dim shapelist As New SaveShapeTypeOption()
            shapelist.SaveAll = True

            ' Save the shapes in the worksheet as images and store them in a list
            Dim images As List(Of Bitmap) = worksheet.SaveShapesToImage(shapelist)

            ' Initialize the index variable for naming the image files
            Dim index As Integer = 0

            ' Iterate through each image in the list
            For Each img As Image In images
                ' Generate a unique file name for each image
                Dim imageFileName As String = "Image_" & index & ".png"

                ' Save the image as a PNG file
                img.Save(imageFileName, ImageFormat.Png)

                ' Increment the index for the next image
                index += 1
            Next img

            ' Dispose the workbook object
            workbook.Dispose()
        End Sub

		Private Sub OutputViewer(ByVal filename As String)
			Try
				Process.Start(filename)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
