Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes

Namespace AdjustArrowPolylinePosition
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Add an elbow line shape to the worksheet and set its properties
            Dim line As XlsLineShape = TryCast(worksheet.TypedLines.AddLine(5, 5, 100, 100, LineShapeType.ElbowLine), Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape)
            line.EndArrowHeadStyle = ShapeArrowStyleType.LineNoArrow
            line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow

            ' Add a geometry adjust value to the line shape and set its formula parameter
            Dim ad As GeomertyAdjustValue = line.ShapeAdjustValues.AddAdjustValue(GeomertyAdjustValueFormulaType.LiteralValue)

            ' Explanation of the following line:
            ' When the parameter value is less than 0, the focus of the line is on the left side of the left point,
            ' when it is equal to 0, the position is the same as the left point,
            ' it is equal to 50 in the middle of the graph,
            ' and when it is equal to 100, it is the same as the right point.
            ad.SetFormulaParameter(New Double() {-50})

            ' Specify the output file name for saving the modified workbook
            Dim result As String = "AdjustArrowPolylinePosition.xlsx"

            ' Save the workbook to the specified output file path in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
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
