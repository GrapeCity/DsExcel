Imports System.Drawing

Namespace Features.Worksheets
    Public Class ConfigureWorksheetView
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Worksheet view settings.
            Dim sheetView As IWorksheetView = worksheet.SheetView
            sheetView.DisplayFormulas = False
            sheetView.DisplayRightToLeft = True
            sheetView.GridlineColor = Color.Red
            sheetView.Zoom = 200
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowScreenshot As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
