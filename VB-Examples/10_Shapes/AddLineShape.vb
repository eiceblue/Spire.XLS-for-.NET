Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddLineShape
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

			'Add shape line1
			Dim line1 As ILineShape = sheet.Lines.AddLine(10, 2, 200, 1, LineShapeType.Line)
			'Set dash style type
			line1.DashStyle = ShapeDashLineStyleType.Solid
			'Set color
			line1.Color = Color.CadetBlue
			'Set weight
			line1.Weight = 2f
			'Set end arrow style type
			line1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow

			'Add shape line2
			Dim line2 As ILineShape = sheet.Lines.AddLine(12, 2, 200, 1, LineShapeType.CurveLine)
			line2.DashStyle = ShapeDashLineStyleType.Dotted
			line2.Color = Color.OrangeRed
			line2.Weight = 2f

			'Add shape line3
			Dim line3 As ILineShape = sheet.Lines.AddLine(14, 2, 200, 1, LineShapeType.ElbowLine)
			line3.DashStyle = ShapeDashLineStyleType.DashDotDot
			line3.Color = Color.Purple
			line3.Weight = 2f

			'Add shape line4
			Dim line4 As ILineShape = sheet.Lines.AddLine(16, 2, 200, 1, LineShapeType.LineInv)
			line4.DashStyle = ShapeDashLineStyleType.Dashed
			line4.Color = Color.Green
			line4.Weight = 2f

			'Save the document
			Dim output As String = "InsertLineShape_out.xlsx"
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
