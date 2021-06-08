Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace AddArrowLineToExcelFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add a Double Arrow and fill the line with solid color.
			Dim line = sheet.TypedLines.AddLine()
			line.Top = 10
			line.Left = 20
			line.Width = 100
			line.Height = 0
			line.Color = Color.Blue
			line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
			line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow

			'Add an Arrow and fill the line with solid color.
			Dim line_1 = sheet.TypedLines.AddLine()
			line_1.Top = 50
			line_1.Left = 30
			line_1.Width = 100
			line_1.Height = 100
			line_1.Color = Color.Red
			line_1.BeginArrowHeadStyle = ShapeArrowStyleType.LineNoArrow
			line_1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow

			'Add an Elbow Arrow Connector.
			Dim line3 As Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape = TryCast(sheet.TypedLines.AddLine(), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
			line3.LineShapeType = LineShapeType.ElbowLine
			line3.Width = 30
			line3.Height = 50
			line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
			line3.Top = 100
			line3.Left = 50

			'Add an Elbow Double-Arrow Connector.
			Dim line2 As Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape = TryCast(sheet.TypedLines.AddLine(), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
			line2.LineShapeType = LineShapeType.ElbowLine
			line2.Width = 50
			line2.Height = 50
			line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
			line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
			line2.Left = 120
			line2.Top = 100

			'Add a Curved Arrow Connector.
			line3 = TryCast(sheet.TypedLines.AddLine(), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
			line3.LineShapeType = LineShapeType.CurveLine
			line3.Width = 30
			line3.Height = 50
			line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
			line3.Top = 100
			line3.Left = 200

			'Add a Curved Double-Arrow Connector.
			line2 = TryCast(sheet.TypedLines.AddLine(), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
			line2.LineShapeType = LineShapeType.CurveLine
			line2.Width = 30
			line2.Height = 50
			line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
			line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
			line2.Left = 250
			line2.Top = 100

			Dim result As String = "Result-AddArrowLineToExcelFile.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the MS Excel file.
			ExcelDocViewer(result)
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
