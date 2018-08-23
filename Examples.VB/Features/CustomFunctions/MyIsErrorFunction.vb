Option Strict On

Namespace Features.CustomFunctions
    Public Class MyIsErrorFunction
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Excel.Workbook.AddCustomFunction(New MyIsErrorFunctionX)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range!A1.Value = CalcError.Num
            worksheet.Range!A2.Value = 100

            worksheet.Range!B1.Formula = "=MyIsError(A1)"
            worksheet.Range!B2.Formula = "=MyIsError(A2)"

            ' Range!B1.Value = True.
            Dim resultB1 = worksheet.Range!B1.Value

            ' Range!B2.Value = False.
            Dim resultB2 = worksheet.Range!B2.Value

            '    Implementation of MyIsErrorFunctionX
            '      
            '    Public Class MyIsErrorFunctionX
            '        Inherits CustomFunction

            '        Public Sub New()
            '            MyBase.New("MyIsError", FunctionValueType.Boolean, {New Parameter(FunctionValueType.Variant)})
            '        End Sub

            '        Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
            '            If TypeOf arguments(0) Is CalcError Then
            '                Dim calcError = DirectCast(arguments(0), CalcError)
            '                Return calcError <> CalcError.None AndAlso
            '                       calcError <> CalcError.GettingData
            '            End If
            '            Return False
            '        End Function
            '    End Class
            '            
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

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Class MyIsErrorFunctionX
            Inherits CustomFunction

            Public Sub New()
                MyBase.New("MyIsError", FunctionValueType.Boolean, {New Parameter(FunctionValueType.Variant)})
            End Sub

            Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
                If TypeOf arguments(0) Is CalcError Then
                    Dim calcError = DirectCast(arguments(0), CalcError)
                    Return calcError <> CalcError.None AndAlso
                           calcError <> CalcError.GettingData
                End If
                Return False
            End Function
        End Class

    End Class
End Namespace
