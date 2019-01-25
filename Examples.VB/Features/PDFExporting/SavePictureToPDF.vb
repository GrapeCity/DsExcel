Imports System.Reflection
Namespace Features.PDFExporting
    Public Class SavePictureToPDF
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.PageSetup.Orientation = PageOrientation.Landscape

            Dim stream As IO.Stream = GetResourceStream("logo.png")
            Dim picture As IShape = worksheet.Shapes.AddPicture(stream, ImageType.PNG, 20, 20, 690, 100)
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
