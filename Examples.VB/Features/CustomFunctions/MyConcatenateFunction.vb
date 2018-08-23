Imports System.Text

Namespace Features.CustomFunctions
    Public Class MyConcatenateFunction
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Excel.Workbook.AddCustomFunction(New MyConcatenateFunctionX)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Formula = "=MyConcatenate(""I"", "" "", ""live"", "" "", ""in"", "" "", ""Xi'an"", ""."")"
            worksheet.Range!A2.Formula = "=MyConcatenate(A1, ""haha."")"

            worksheet.Range!B1.Value = 12
            worksheet.Range!B2.Value = 34
            worksheet.Range!B3.Formula = "=MyConcatenate(B1, B2)"

            worksheet.Range("M5:N5").FormulaArray = "=CONCATENATE({""aa"",""bb""}, 12, 34)"

            ' "I live in Xi'an."
            Dim resultA1 = worksheet.Range!A1.Value

            ' "I live in Xi'an.haha."
            Dim resultA2 = worksheet.Range!A2.Value

            ' "1234"
            Dim resultB3 = worksheet.Range!B3.Value

            ' "aa1234"
            Dim resultM5 = worksheet.Range!M5.Value

            ' "bb1234"
            Dim resultN5 = worksheet.Range!N5.Value

            '    Implementation of MyAddFunctionX
            '    
            '    Public Class MyConcatenateFunctionX
            '        Inherits CustomFunction

            '        Public Sub New()
            '            MyBase.New("MyConcatenate", FunctionValueType.Text, CreateParameters())
            '        End Sub

            '        Private Shared Function CreateParameters() As Parameter()
            '            Dim parameters(253) As Parameter
            '            For i As Integer = 0 To 253
            '                parameters(i) = New Parameter(FunctionValueType.Variant)
            '            Next

            '            Return parameters
            '        End Function

            '        Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
            '            Dim sb As New StringBuilder

            '            Dim result As String = String.Empty
            '            For Each argument In arguments
            '                If TypeOf argument Is CalcError Then
            '                    Return argument
            '                End If
            '                If TypeOf argument Is String OrElse TypeOf argument Is Double Then
            '                    sb.Append(argument)
            '                End If
            '            Next

            '            Return sb.ToString()
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

        Public Class MyConcatenateFunctionX
            Inherits CustomFunction

            Public Sub New()
                MyBase.New("MyConcatenate", FunctionValueType.Text, CreateParameters())
            End Sub

            Private Shared Function CreateParameters() As Parameter()
                Dim parameters(253) As Parameter
                For i As Integer = 0 To 253
                    parameters(i) = New Parameter(FunctionValueType.Variant)
                Next

                Return parameters
            End Function

            Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
                Dim sb As New StringBuilder

                Dim result As String = String.Empty
                For Each argument In arguments
                    If TypeOf argument Is CalcError Then
                        Return argument
                    End If
                    If TypeOf argument Is String OrElse TypeOf argument Is Double Then
                        sb.Append(argument)
                    End If
                Next

                Return sb.ToString()
            End Function
        End Class
    End Class
End Namespace
