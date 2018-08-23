Namespace Features.Formatting.Protection
    Public Class SetRangeFormulaHidden
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range!B1.Formula = "=A1"

            'config range B1's FormulaHidden property.
            worksheet.Range!B1.FormulaHidden = True

            'protect worksheet, range B1's formula will not show in exported xlsx file.
            worksheet.Protection = True
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
