
Namespace Features.PDFExporting
    Public Class PDFToExcel
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim fs As IO.Stream = GetResourceStream("TimeSheet.pdf")

            Dim doc1 As GrapeCity.Documents.Pdf.GcPdfDocument = New GrapeCity.Documents.Pdf.GcPdfDocument()
            doc1.Load(fs)
            Dim tmap As GrapeCity.Documents.Pdf.TextMap.ITextMap = doc1.Pages(0).GetTextMap()
            Dim tx0 As Single = 2.1F, ty0 As Single = 3.37F, tx1 As Single = 3.1F, ty1 As Single = 3.5F
            Dim htiFrom As GrapeCity.Documents.Pdf.TextMap.HitTestInfo = tmap.HitTest(tx0 * 72, ty0 * 72)
            Dim htiTo As GrapeCity.Documents.Pdf.TextMap.HitTestInfo = tmap.HitTest(ty0 * 72, ty1 * 72)
            Dim range1 As GrapeCity.Documents.Pdf.TextMap.TextMapFragment = Nothing, text1 As String = Nothing
            tmap.GetFragment(htiFrom.Pos, htiTo.Pos, range1, text1)
            worksheet.Range("A1").Value = "List of all texts found in the TimeSheet.pdf"
            Dim range As GrapeCity.Documents.Pdf.TextMap.TextMapFragment = Nothing, text As String = Nothing
            tmap.GetFragment(range, text)
            Dim row As Integer = 1
            worksheet.Columns(0).ColumnWidth = 20

            For Each tlf As GrapeCity.Documents.Pdf.TextMap.TextLineFragment In range
                Dim coords = tmap.GetCoords(tlf)
                worksheet.Range(row, 0).Value = $"Text at ({coords.B.X / 72}"",{coords.B.Y / 72}""):"
                worksheet.Range(row, 1).Value = tmap.GetText(tlf)
                row += 1
            Next
        End Sub
        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
