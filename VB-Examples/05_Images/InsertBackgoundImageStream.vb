Imports Spire.Xls
Imports System
Imports System.IO
Imports System.Windows.Forms

Namespace InsertBackgoundImageStream
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Create a workbook
            Dim workbook As Workbook = New Workbook()

            'Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

            'Get the first worksheet.
            Dim sheet As Worksheet = workbook.Worksheets[0]

            ' Open the image as a stream
            Dim image As Stream = File.OpenRead(@"..\..\..\..\..\..\Data\Background.emf")

            'Set the image stream to be background image of the worksheet.
            sheet.PageSetup.BackgoundImageStream = image

            ' Specify the output file name for the EXCEL result
            Dim result As String = "InsertBackgoundImageStream_out.xlsx"

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
