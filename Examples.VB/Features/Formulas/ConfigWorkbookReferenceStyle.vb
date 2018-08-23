Namespace Features.Formulas
    Public Class ConfigWorkbookReferenceStyle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'set workbook's reference style to R1C1. exported xlsx file will be R1C1 style.
            workbook.ReferenceStyle = ReferenceStyle.R1C1
        End Sub
    End Class
End Namespace
