Imports System.Drawing

Namespace Features.Worksheets
    Public Class ConfigWorksheet
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            'Set worksheet tab color.
            worksheet.TabColor = Color.Green

            'Set worksheet default row height.
            worksheet.StandardHeight = 20

            'Set worksheet default column width.
            worksheet.StandardWidth = 50

            'Split worksheet to panes.
            worksheet.SplitPanes(worksheet.Range!B3.Row, worksheet.Range!B3.Column)
            Dim worksheet1 As IWorksheet = workbook.Worksheets.Add()

            'Hide worksheet.
            worksheet1.Visible = Visibility.Hidden
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
        Public Overrides ReadOnly Property IsUpdate As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
