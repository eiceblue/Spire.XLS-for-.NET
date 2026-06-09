Imports Spire.Xls
Imports System
Imports System.Windows.Forms

Namespace XMATCHFunction
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new workbook
            Dim workbook As Workbook = New Workbook()

            ' Load an existing workbook with a pivot table from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\XMATCHFunction")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets[0]

            ' Set the formula for cell C4 to use XMATCH function
            sheet.Range["C4"].Formula = "=XMATCH(\"Lili\", A2:A5)"

            'Calculate all cells
            workbook.CalculateAllValue()

            ' Specify the output file name for the result
            Dim result As String = "XMATCHFunction_result.xlsx"

            ' Save the modified workbook to the specified file using Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)

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
