Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System
Imports System.IO
Imports System.Windows.Forms

Namespace InsertEmbedCheckBox
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new workbook object
            Dim workbook As Workbook = New Workbook()

            ' Get the reference to the first sheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets[0]

            ' Get the cell range at position A1
            Dim range As XlsRange = sheet.Range["A1"]

            ' Insert an embedded checkbox into the specified range
            range.InsertEmbedCheckBox()

            ' Set the checkbox state to checked (true = checked, false = unchecked)
            range.SetEmbedCheckBoxCheckState(True)

            ' Specify the output file name for the EXCEL result
            Dim result As String = "InsertEmbedCheckBox_out.xlsx"

            'Save the Excel file
            workbook.SaveToFile(result, FileFormat.Version2013)

            ' Dispose of the workbook object to release resources
            workbook.Dispose()

            'View the document
            FileViewer(result)
        End Sub

        Private Sub FileViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

        Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
            Close()
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace
