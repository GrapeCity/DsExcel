Namespace Features.Workbook
    Public Class ImportXlsmToWorkbook
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)

            'GcExcel supports open xlsm file
            workbook.Open(Me.CurrentDirectory & "macros.xlsm")
            'Macros can be preserved after saving
            workbook.Save(Me.CurrentDirectory & "macros-exported.xlsm")
        End Sub
        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property IsUpdate As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
