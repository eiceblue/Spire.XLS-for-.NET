Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.Security.Cryptography

Namespace ShapeToImageOptions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

			' Create a new Workbook object
			Dim workbook As New Workbook()

			' Load the workbook from the specified file path
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Shape.xlsx")

			' Get the first worksheet from the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Convert shapes to images
			Dim shapelist As New SaveShapeTypeOption()

			' Set the option to save all shapes in the worksheet to images
			shapelist.SaveAll = True

			' Save the shapes in the worksheet as images and store them in a dictionary
			Dim images As Dictionary(Of IShape, Bitmap) = sheet.SaveAndGetShapesToImage(shapelist)

			' Iterate over each shape-image pair in the dictionary
			For Each pair As KeyValuePair(Of IShape, Bitmap) In images
				' Get the shape and image from the pair
				Dim shape As IShape = pair.Key
				Dim bitmap As Bitmap = pair.Value

				' Generate a unique image file name based on shape properties
				Dim imageFileName As String = shape.Name & "_" & shape.Height & "_" & shape.Width & "_" & shape.ShapeType & ".png"

				' Save the bitmap as an image file with the generated name
				bitmap.Save(imageFileName)

				OutputViewer(imageFileName)
			Next pair

			' Close Workbook
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
