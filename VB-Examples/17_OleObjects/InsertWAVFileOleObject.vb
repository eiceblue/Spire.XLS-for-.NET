Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.Drawing

Namespace InsertWavFileOLEObject
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add an OLE object to the worksheet, with the specified WAV file path and an image as an icon
            Dim oleObject As IOleObject = sheet.OleObjects.Add("..\..\..\..\..\..\Data\WAVFileSample.wav", Image.FromFile("..\..\..\..\..\..\Data\SpireXls.png"), OleLinkType.Embed)

            ' Set the location of the OLE object on the worksheet
            oleObject.Location = sheet.Range("B4")

            ' Set the type of the OLE object as Package
            oleObject.ObjectType = OleObjectType.Package

            ' Specify the filename for the resulting Excel file
            Dim result As String = "result.xlsx"

            ' Save the workbook to the specified filename in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer(result)

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
