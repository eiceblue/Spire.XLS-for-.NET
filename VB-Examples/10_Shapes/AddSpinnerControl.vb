Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddSpinnerControl
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of a Workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text in cell C11 as "Value:" and make it bold
            sheet.Range("C11").Text = "Value:"
            sheet.Range("C11").Style.Font.IsBold = True

            ' Set the value in cell C12 as 0
            sheet.Range("C12").Value2 = 0

            ' Add a spinner control to the worksheet at position (12, 4) with size 20x20
            Dim spinner As ISpinnerShape = sheet.SpinnerShapes.AddSpinner(12, 4, 20, 20)

            ' Link the spinner control to cell C12
            spinner.LinkedCell = sheet.Range("C12")

            ' Set the minimum and maximum values for the spinner
            spinner.Min = 0
            spinner.Max = 100

            ' Specify the incremental change when the spinner is clicked
            spinner.IncrementalChange = 5

            ' Enable 3D shading for the spinner control
            spinner.Display3DShading = True

            ' Save the modified workbook to a new file called "AddSpinnerControl_out.xlsx"
            Dim output As String = "AddSpinnerControl_out.xlsx"
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
