Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddLabelControl
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Excel workbook object.
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            ' Access the first worksheet of the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create and add a label control to the worksheet at the specified position and size.
            Dim label As ILabelShape = sheet.LabelShapes.AddLabel(10, 2, 30, 200)
            'Set the text content of the label control.
            label.Text = "This is a Label Control"

            ' Specify the file name for the resulting Excel file.
            Dim output As String = "InsertLabelControl_out.xlsx"

            'Save the workbook to a file.
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
