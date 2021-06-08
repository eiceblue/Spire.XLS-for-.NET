Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddOvalShape
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add oval shape1
			Dim ovalShape1 As IOvalShape = sheet.OvalShapes.AddOval(11, 2, 100, 100)
			ovalShape1.Line.Weight = 0
			'Fill shape with solid color
			ovalShape1.Fill.FillType = ShapeFillType.SolidColor
			ovalShape1.Fill.ForeColor = Color.DarkCyan

			'Add oval shape2
			Dim ovalShape2 As IOvalShape = sheet.OvalShapes.AddOval(11, 5, 100, 100)
			ovalShape2.Line.Weight = 1
			'Fill shape with picture
			ovalShape2.Line.DashStyle = ShapeDashLineStyleType.Solid
			ovalShape2.Fill.CustomPicture("..\..\..\..\..\..\Data\logo.png")

			'Save the document
			Dim output As String = "AddOvalShape_out.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

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
