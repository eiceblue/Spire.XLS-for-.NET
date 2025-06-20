Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports Spire.Xls

Namespace ToOfficeOpenXML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object.
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text of cell A1 in the worksheet to "Hello World".
            sheet.Range("A1").Text = "Hello World"

            ' Set the fill color of cell B1 to 25% Gray.
            sheet.Range("B1").Style.KnownColor = ExcelColors.Gray25Percent

            ' Set the fill color of cell C1 to Gold.
            sheet.Range("C1").Style.KnownColor = ExcelColors.Gold

            ' Save the workbook as an XML file with the specified file name.
            workbook.SaveAsXml("sample.xml")
            ' Release the resources used by the workbook
            workbook.Dispose()

            Process.Start(Path.Combine(Application.StartupPath,"Sample.xml"))
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
