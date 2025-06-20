Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.IO

Namespace ExtractOLEObjects
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExtractOle2.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Check if the worksheet has any OLE objects
            If sheet.HasOleObjects Then
                ' Iterate through each OLE object in the worksheet
                For i As Integer = 0 To sheet.OleObjects.Count - 1
                    ' Get the current OLE object
                    Dim [Object] = sheet.OleObjects(i)

                    ' Get the type of the current OLE object
                    Dim type As OleObjectType = sheet.OleObjects(i).ObjectType

                    ' Process the OLE object based on its type
                    Select Case type
                        Case OleObjectType.WordDocument
                            ' Save the OLE data as a Word document file
                            File.WriteAllBytes("Ole.docx", [Object].OleData)

                        Case OleObjectType.PowerPointSlide
                            ' Save the OLE data as a PowerPoint slide file
                            File.WriteAllBytes("Ole.pptx", [Object].OleData)

                        Case OleObjectType.AdobeAcrobatDocument
                            ' Save the OLE data as a PDF file
                            File.WriteAllBytes("Ole.pdf", [Object].OleData)
                    End Select
                Next i
            End If

            ' Release the resources used by the workbook
            workbook.Dispose()
            MessageBox.Show("Completed!")
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
