Imports Spire.Xls
Imports System
Imports System.IO
Imports System.Windows.Forms

Namespace SpecificCellsToHtml
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Create a workbook
            Dim workbook As Workbook = New Workbook()

            'Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConversionSample1.xlsx")

            'Get the first worksheet in Excel file
            Dim sheet As Worksheet = workbook.Worksheets[0]

            ' Get the specific cell range A1:B3 from the worksheet
            Dim cell As CellRange = sheet.Range["A1:E7"]

            ' Extract the HTML representation of the selected cell range
            Dim html As String = cell.HtmlString

            ' Specify the output file name for the HTML result
            Dim result As String = "SpecificCellsToHtml_out.html"

            ' Write the HTML content to the output file
            File.WriteAllText(result, html)

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
