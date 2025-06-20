Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace FillPicAsTextureInShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\TillPicAsTextureInShape.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first preset geometry shape from the worksheet
            Dim shape As IPrstGeomShape = sheet.PrstGeomShapes(0)

            ' Set the fill type of the shape to Texture
            shape.Fill.FillType = ShapeFillType.Texture

            ' Set the custom texture for the shape using the specified image file path
            shape.Fill.CustomTexture("..\..\..\..\..\..\Data\logo.png")
            shape.Fill.Tile = True

            ' Specify the output file path for saving the modified workbook
            Dim output As String = "TillPicAsTextureInShape_out.xlsx"

            ' Save the workbook to the specified output file path in Excel 2013 format
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
