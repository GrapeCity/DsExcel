Namespace Features.CustomFunctions
    Public Class MyAddFunction
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Excel.Workbook.AddCustomFunction(New MyAddFunctionX)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = 1
            worksheet.Range!B1.Value = 2
            worksheet.Range!C1.Formula = "=MyAdd(A1, B1)"

            ' Range!C1.Value = 3.
            Dim result = worksheet.Range!C1.Value

            worksheet.Range("E1:F2").Value = New Object(,) {
                {1, 3},
                {2, 4}
            }

            ' Custom function use in array formula.
            worksheet.Range("G1:G2").FormulaArray = "=MyAdd(E1:E2, F1:F2)"

            ' Range!G1.Value = 4, Range!G2.Value = 6.
            Dim resultG1 = worksheet.Range!G1.Value
            Dim resultG2 = worksheet.Range!G2.Value

            '    Implementation of MyAddFunctionX
            '    
            '    Public Class MyAddFunctionX
            '        Inherits CustomFunction

            '        Public Sub New()
            '            MyBase.New("MyAdd", FunctionValueType.Number, {
            '                New Parameter(FunctionValueType.Number),
            '                New Parameter(FunctionValueType.Number)
            '            })

            '        End Sub

            '        Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
            '            Return CDbl(arguments(0)) + CDbl(arguments(1))
            '        End Function
            '    End Class
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

        Public Class MyAddFunctionX
            Inherits CustomFunction

            Public Sub New()
                MyBase.New("MyAdd", FunctionValueType.Number, {
                    New Parameter(FunctionValueType.Number),
                    New Parameter(FunctionValueType.Number)
                })

            End Sub

            Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
                Return CDbl(arguments(0)) + CDbl(arguments(1))
            End Function
        End Class

    End Class
End Namespace
