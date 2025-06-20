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
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a line shape to the worksheet and set its properties
            Dim line = sheet.TypedLines.AddLine()
            line.Top = 10
            line.Left = 20
            line.Width = 100
            line.Height = 0
            line.Color = Color.Blue
            line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
            line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow

            ' Add another line shape to the worksheet and set its properties
            Dim line_1 = sheet.TypedLines.AddLine()
            line_1.Top = 50
            line_1.Left = 30
            line_1.Width = 100
            line_1.Height = 100
            line_1.Color = Color.Red
            line_1.BeginArrowHeadStyle = ShapeArrowStyleType.LineNoArrow
            line_1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow

            ' Add a third line shape of a specific type to the worksheet and set its properties
            Dim line3 As Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape = TryCast(sheet.TypedLines.AddLine(), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
            line3.LineShapeType = LineShapeType.ElbowLine
            line3.Width = 30
            line3.Height = 50
            line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
            line3.Top = 100
            line3.Left = 50

            ' Add a fourth line shape of a specific type to the worksheet and set its properties
            Dim line2 As Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape = TryCast(sheet.TypedLines.AddLine(), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
            line2.LineShapeType = LineShapeType.ElbowLine
            line2.Width = 50
            line2.Height = 50
            line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow
            line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
            line2.Left = 120
            line2.Top = 100

            ' Add a fifth line shape of a specific type to the worksheet and set its properties
            line3 = TryCast(sheet.TypedLines.AddLine(), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
            line3.LineShapeType = LineShapeType.CurveLine
            line3.Width = 30
            line3.Height = 50
            line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
            line3.Top = 100
            line3.Left = 200

            ' Add a sixth line shape of a specific type to the worksheet and set its properties
            line2 = TryCast(sheet.TypedLines.AddLine(), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
            line2.LineShapeType = LineShapeType.CurveLine
            line2.Width = 30
            line2.Height = 50
            line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
            line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen
            line2.Left = 250
            line2.Top = 100

            ' Specify the output file name for saving the modified workbook
            Dim result As String = "Result-AddArrowLineToExcelFile.xlsx"

            ' Save the workbook to the specified output file path in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
