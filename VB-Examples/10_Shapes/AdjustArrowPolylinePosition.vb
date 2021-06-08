Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace AdjustArrowPolylinePosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Draw an elbow arrow
			Dim line As XlsLineShape = TryCast(worksheet.TypedLines.AddLine(5, 5, 100, 100, LineShapeType.ElbowLine), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
			line.EndArrowHeadStyle = ShapeArrowStyleType.LineNoArrow
			line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow
			Dim ad As GeomertyAdjustValue = line.ShapeAdjustValues.AddAdjustValue(GeomertyAdjustValueFormulaType.LiteralValue)

			'When the parameter value is less than 0, the focus of the line is on the left side of the left point, when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, and when it is equal to 100, it is the same as the right point.
			ad.SetFormulaParameter(New Double() {-50})

			'Save to file
			Dim result As String = "AdjustArrowPolylinePosition.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
			FileViewer(result)
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
