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
            ' Declare and initialize a new Workbook object
            Dim workbook As New Workbook()

            ' Load an Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add an oval shape to the worksheet at position (11, 2) with dimensions 100x100
            Dim ovalShape1 As IOvalShape = sheet.OvalShapes.AddOval(11, 2, 100, 100)
            ovalShape1.Line.Weight = 0 ' Set the weight of the line to 0

            ' Set the fill type of ovalShape1 to solid color and set the foreground color to DarkCyan
            ovalShape1.Fill.FillType = ShapeFillType.SolidColor
            ovalShape1.Fill.ForeColor = Color.DarkCyan

            ' Add another oval shape to the worksheet at position (11, 5) with dimensions 100x100
            Dim ovalShape2 As IOvalShape = sheet.OvalShapes.AddOval(11, 5, 100, 100)
            ovalShape2.Line.Weight = 1 ' Set the weight of the line to 1

            ' Set the dash style of the line for ovalShape2 to solid and set a custom picture fill
            ovalShape2.Line.DashStyle = ShapeDashLineStyleType.Solid
            ovalShape2.Fill.CustomPicture("..\..\..\..\..\..\Data\logo.png")

            ' Specify the output file name
            Dim output As String = "AddOvalShape_out.xlsx"

            ' Save the modified workbook to the specified file in Excel 2013 format
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
