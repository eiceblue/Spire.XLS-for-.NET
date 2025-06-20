Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace InsertOLEObjects
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim ws As Worksheet = workbook.Worksheets(0)

            ' Set the text in cell A1 of the worksheet
            ws.Range("A1").Text = "Here is an OLE Object."

            ' Specify the file path for the Excel file to be inserted as an OLE object
            Dim xlsFile As String = "..\..\..\..\..\..\Data\InsertOLEObjects.xls"

            ' Generate an image from the Excel file
            Dim image As Image = GenerateImage(xlsFile)

            ' Add an OLE object to the worksheet with the specified Excel file and image
            Dim oleObject As IOleObject = ws.OleObjects.Add(xlsFile, image, OleLinkType.Embed)

            ' Set the location of the OLE object on the worksheet (B4)
            oleObject.Location = ws.Range("B4")

            ' Set the object type of the OLE object to indicate it is an Excel worksheet
            oleObject.ObjectType = OleObjectType.ExcelWorksheet

            ' Specify the output file name for the saved workbook
            Dim result As String = "InsertOLEObjects_result.xlsx"

            ' Save the workbook to the specified file path using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()
            ExcelDocViewer(result)
		End Sub
		Private Function GenerateImage(ByVal fileName As String) As Image
            ' Create a new instance of Workbook
            Dim book As New Workbook()

            ' Load the workbook from the specified file path
            book.LoadFromFile(fileName)

            ' Configure the page setup to remove margins
            book.Worksheets(0).PageSetup.LeftMargin = 0
            book.Worksheets(0).PageSetup.RightMargin = 0
            book.Worksheets(0).PageSetup.TopMargin = 0
            book.Worksheets(0).PageSetup.BottomMargin = 0

            ' Convert the worksheet to an image, specifying the range of cells to capture
            Return book.Worksheets(0).ToImage(1, 1, 19, 5)
        End Function
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
