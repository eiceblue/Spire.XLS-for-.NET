Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes
Imports System.Drawing.Imaging

Namespace ShapeToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ShapeToImage.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet1 As Worksheet = workbook.Worksheets(0)

            ' Get the first PrstGeomShape in the worksheet and cast it to XlsShape type
            Dim shape As XlsShape = TryCast(sheet1.PrstGeomShapes(0), XlsShape)

            ' Save the shape as an image
            Dim img As Image = shape.SaveToImage()

            ' Save the image to a file named "ShapeToImage.png" in PNG format
            img.Save("ShapeToImage.png", ImageFormat.Png)
            ' Release the resources used by the workbook
            workbook.Dispose()
            FileViewer("ShapeToImage.png")
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
