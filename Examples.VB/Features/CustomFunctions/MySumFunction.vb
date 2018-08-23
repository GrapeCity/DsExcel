Namespace Features.CustomFunctions
    Public Class MySumFunction
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Excel.Workbook.AddCustomFunction(New MySumFunctionX)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = 1
            worksheet.Range!B1.Value = 2
            worksheet.Range!C1.Formula = "=MySum(A1:B1, 2, {3,4})"

            ' Range!C1.Value = 12.
            Dim result = worksheet.Range!C1.Value

            '    Implementation of MySumFunctionX
            '
            '    Public Class MySumFunctionX
            '        Inherits CustomFunction

            '        Public Sub New()
            '            MyBase.New("MYSUM", FunctionValueType.Number, CreateParameters())
            '        End Sub

            '        Private Shared Function CreateParameters() As Parameter()
            '            Dim parameters(253) As Parameter
            '            For i As Integer = 0 To 253
            '                parameters(i) = New Parameter(FunctionValueType.Object)
            '            Next
            '            Return parameters
            '        End Function

            '        Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
            '            Dim sum As Double = 0R
            '            For Each argument In arguments
            '                For Each item In Enumerate(argument)
            '                    If TypeOf item Is CalcError Then
            '                        Return item
            '                    End If
            '                    If TypeOf item Is Double Then
            '                        sum += CDbl(item)
            '                    End If
            '                Next item
            '            Next argument
            '            Return sum
            '        End Function

            '        Private Shared Iterator Function Enumerate(obj As Object) As IEnumerable(Of Object)
            '            Dim objects = TryCast(obj, IEnumerable(Of Object))
            '            If objects IsNot Nothing Then
            '                For Each item In objects
            '                    For Each item2 In Enumerate(item)
            '                        Yield item2
            '                    Next item2
            '                Next item
            '            Else
            '                Dim array = TryCast(obj, Object(,))
            '                If array IsNot Nothing Then
            '                    Dim rowCount As Integer = array.GetLength(0)
            '                    Dim colCount As Integer = array.GetLength(1)
            '                    For i As Integer = 0 To rowCount - 1
            '                        For j As Integer = 0 To colCount - 1
            '                            Yield array(i, j)
            '                        Next j
            '                    Next i
            '                End If
            '            End If
            '            If obj IsNot Nothing Then
            '                Yield obj
            '            End If
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

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Class MySumFunctionX
            Inherits CustomFunction

            Public Sub New()
                MyBase.New("MYSUM", FunctionValueType.Number, CreateParameters())
            End Sub

            Private Shared Function CreateParameters() As Parameter()
                Dim parameters(253) As Parameter
                For i As Integer = 0 To 253
                    parameters(i) = New Parameter(FunctionValueType.Object)
                Next
                Return parameters
            End Function

            Public Overrides Function Evaluate(arguments As Object(), context As ICalcContext) As Object
                Dim sum As Double = 0R
                For Each argument In arguments
                    For Each item In Enumerate(argument)
                        If TypeOf item Is CalcError Then
                            Return item
                        End If
                        If TypeOf item Is Double Then
                            sum += CDbl(item)
                        End If
                    Next item
                Next argument

                Return sum
            End Function

            Private Shared Iterator Function Enumerate(obj As Object) As IEnumerable(Of Object)
                Dim objects = TryCast(obj, IEnumerable(Of Object))
                If objects IsNot Nothing Then
                    For Each item In objects
                        For Each item2 In Enumerate(item)
                            Yield item2
                        Next item2
                    Next item
                Else
                    Dim array = TryCast(obj, Object(,))
                    If array IsNot Nothing Then
                        Dim rowCount As Integer = array.GetLength(0)
                        Dim colCount As Integer = array.GetLength(1)
                        For i As Integer = 0 To rowCount - 1
                            For j As Integer = 0 To colCount - 1
                                Yield array(i, j)
                            Next j
                        Next i
                    End If
                End If

                If obj IsNot Nothing Then
                    Yield obj
                End If
            End Function
        End Class
    End Class

End Namespace
