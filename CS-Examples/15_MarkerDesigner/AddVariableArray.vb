Imports Spire.Xls

Namespace AddVariableArray

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the value of cell A1 to "&=Array"
            sheet.Range("A1").Value = "&=Array"

            ' Add an array marker named "Array" with the specified values to the MarkerDesigner
            workbook.MarkerDesigner.AddArray("Array", New String() {"Spire.Xls", "Spire.Doc", "Spire.PDF", "Spire.Presentation", "Spire.Email"})

            ' Apply the marker designer to replace the markers with actual values
            workbook.MarkerDesigner.Apply()

            ' Calculate all the formulas in the workbook
            workbook.CalculateAllValue()

            ' Autofit the rows and columns of the allocated range in the worksheet
            sheet.AllocatedRange.AutoFitRows()
            sheet.AllocatedRange.AutoFitColumns()

            ' Define the output file name as "AddVariableArray.xlsx"
            Dim output As String = "AddVariableArray.xlsx"

            ' Save the modified workbook to the specified file path using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
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
