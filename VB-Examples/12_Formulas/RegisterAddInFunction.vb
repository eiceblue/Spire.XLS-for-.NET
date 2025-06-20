Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace RegisterAddInFunction
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Define the input file path as a relative path to a Test.xlam add-in
            Dim input As String = "..\..\..\..\..\..\Data\Test.xlam"

            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Add the TEST_UDF function from the specified input add-in to the workbook's AddInFunctions collection
            workbook.AddInFunctions.Add(input, "TEST_UDF")
            ' Add the TEST_UDF1 function from the specified input add-in to the workbook's AddInFunctions collection
            workbook.AddInFunctions.Add(input, "TEST_UDF1")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the formula of cell A1 in the sheet to "=TEST_UDF()"
            sheet.Range("A1").Formula = "=TEST_UDF()"

            ' Set the formula of cell A2 in the sheet to "=TEST_UDF1()"
            sheet.Range("A2").Formula = "=TEST_UDF1()"

            ' Specify the file name for saving the workbook as "RegisterAddInFunction_result.xlsx"
            Dim result As String = "RegisterAddInFunction_result.xlsx"

            ' Save the workbook to a file with the specified name and Excel version 2010
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
