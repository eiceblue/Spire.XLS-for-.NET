Imports Spire.Xls
Imports Spire.Xls.Core.MergeSpreadsheet.Collections
Imports Spire.Xls.Core
Imports System.ComponentModel
Imports System.Text

Namespace GroupShapes
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

            ' Add a rounded rectangle shape to the worksheet and set its properties
            Dim shape1 As IPrstGeomShape = worksheet.PrstGeomShapes.AddPrstGeomShape(1, 3, 50, 50, PrstGeomShapeType.RoundRect)

            ' Add a triangle shape to the worksheet and set its properties
            Dim shape2 As IPrstGeomShape = worksheet.PrstGeomShapes.AddPrstGeomShape(5, 3, 50, 50, PrstGeomShapeType.Triangle)

            ' Get the group shape collection from the worksheet
            Dim groupShapeCollection As GroupShapeCollection = worksheet.GroupShapeCollection

            ' Group the specified shapes together
            groupShapeCollection.Group(New Spire.Xls.Core.IShape() {shape1, shape2})

            ' Specify the output file name for saving the modified workbook
            Dim result As String = "GroupShape.xlsx"

            ' Save the workbook to the specified output file path in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)
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
