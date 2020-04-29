Namespace Features.CustomFunctions
    Public Class MyIsMergedRangeFunction
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Excel.Workbook.AddCustomFunction(New MyIsMergedRangeFunctionX)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1:B2").Merge()

            worksheet.Range!C1.Formula = "=MyIsMergedRange(A1)"
            worksheet.Range!C2.Formula = "=MyIsMergedRange(H2)"

            ' A1 is a merged cell, Range!C1.Value = True.
            Dim resultC1 = worksheet.Range!C1.Value

            ' H2 is not a merged cell, Range!C2.Value = False.
            Dim resultC2 = worksheet.Range!C2.Value

            '    Implementation of MyIsErrorFunctionX
            '    
            '    Public Class MyIsMergedRangeFunctionX
            '        Inherits CustomFunction

            '        Public Sub New()
            '            MyBase.New("MyIsMergedRange", FunctionValueType.Boolean, {New Parameter(FunctionValueType.Object, True)})
            '        End Sub

            '        Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
            '            Dim calcReference = TryCast(arguments(0), CalcReference)
            '            If calcReference IsNot Nothing Then
            '                For Each range In calcReference.GetRanges
            '                    Return range.MergeCells
            '                Next
            '            End If
            '            Return False
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

        Public Class MyIsMergedRangeFunctionX
            Inherits CustomFunction

            Public Sub New()
                MyBase.New("MyIsMergedRange", FunctionValueType.Boolean, {New Parameter(FunctionValueType.Object, True)})
            End Sub

            Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
                Dim calcReference = TryCast(arguments(0), CalcReference)
                If calcReference IsNot Nothing Then
                    For Each range In calcReference.GetRanges
                        Return range.MergeCells
                    Next
                End If
                Return False
            End Function
        End Class

    End Class
End Namespace
