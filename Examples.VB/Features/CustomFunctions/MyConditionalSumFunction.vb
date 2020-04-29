Namespace Features.CustomFunctions
    Public Class MyConditionalSumFunction
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Excel.Workbook.AddCustomFunction(New MyConditionalSumFunctionX)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A1:A10").Value = New Object(,) {
                {1}, {2}, {3}, {4}, {5},
                {6}, {7}, {8}, {9}, {10}
            }

            Dim formatConditions As IFormatConditions = worksheet.Range("A1:A10").FormatConditions
            Dim cellValueRule = formatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Greater, 5)
            DirectCast(cellValueRule, IFormatCondition).Interior.Color = System.Drawing.Color.Red

            ' Sum cells value which display format interior color are red.
            worksheet.Range!C1.Formula = "=MyConditionalSum(A1:A10)"

            ' Range!C1.Value = 40.
            Dim result = worksheet.Range!C1.Value

            '    Implementation of MyConditionalSumFunctionX
            '    
            '    Public Class MyConditionalSumFunctionX
            '        Inherits CustomFunction

            '        Public Sub New()
            '            MyBase.New("MyConditionalSum", FunctionValueType.Number, CreateParameters())
            '        End Sub

            '        Private Shared Function CreateParameters() As Parameter()
            '            Dim parameters(253) As Parameter
            '            For i As Integer = 0 To 253
            '                parameters(i) = New Parameter(FunctionValueType.Object, True)
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
            '            Dim enumerableObjects = TryCast(obj, IEnumerable(Of Object))
            '            If enumerableObjects IsNot Nothing Then
            '                For Each item In enumerableObjects
            '                    For Each item2 In Enumerate(item)
            '                        Yield item2
            '                    Next item2
            '                Next item
            '            Else
            '                Dim objectArray2D = TryCast(obj, Object(,))
            '                If objectArray2D IsNot Nothing Then
            '                    Dim rowCount As Integer = objectArray2D.GetLength(0)
            '                    Dim colCount As Integer = objectArray2D.GetLength(1)
            '                    For i As Integer = 0 To rowCount - 1
            '                        For j As Integer = 0 To colCount - 1
            '                            Yield objectArray2D(i, j)
            '                        Next j
            '                    Next i
            '                Else
            '                    Dim calcReferenceValue = TryCast(obj, CalcReference)
            '                    If calcReferenceValue IsNot Nothing Then
            '                        For Each item In Enumerate(calcReferenceValue)
            '                            Yield item
            '                        Next
            '                    End If
            '                End If
            '            End If

            '            Yield obj
            '        End Function

            '        Private Shared Iterator Function Enumerate(reference As CalcReference) As IEnumerable(Of Object)
            '            For Each range In reference.GetRanges()
            '                Dim rowCount As Integer = range.Rows.Count
            '                Dim colCount As Integer = range.Columns.Count
            '                For i As Integer = 0 To rowCount - 1
            '                    For j As Integer = 0 To colCount - 1
            '                        If range.Cells(i, j).DisplayFormat.Interior.Color = System.Drawing.Color.Red Then
            '                            Yield range.Cells(i, j).Value
            '                        End If
            '                    Next j
            '                Next i
            '            Next range
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

        Public Class MyConditionalSumFunctionX
            Inherits CustomFunction

            Public Sub New()
                MyBase.New("MyConditionalSum", FunctionValueType.Number, CreateParameters())
            End Sub

            Private Shared Function CreateParameters() As Parameter()
                Dim parameters(253) As Parameter
                For i As Integer = 0 To 253
                    parameters(i) = New Parameter(FunctionValueType.Object, True)
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
                Dim enumerableObjects = TryCast(obj, IEnumerable(Of Object))
                If enumerableObjects IsNot Nothing Then
                    For Each item In enumerableObjects
                        For Each item2 In Enumerate(item)
                            Yield item2
                        Next item2
                    Next item
                Else
                    Dim objectArray2D = TryCast(obj, Object(,))
                    If objectArray2D IsNot Nothing Then
                        Dim rowCount As Integer = objectArray2D.GetLength(0)
                        Dim colCount As Integer = objectArray2D.GetLength(1)
                        For i As Integer = 0 To rowCount - 1
                            For j As Integer = 0 To colCount - 1
                                Yield objectArray2D(i, j)
                            Next j
                        Next i
                    Else
                        Dim calcReferenceValue = TryCast(obj, CalcReference)
                        If calcReferenceValue IsNot Nothing Then
                            For Each item In Enumerate(calcReferenceValue)
                                Yield item
                            Next
                        End If
                    End If
                End If

                Yield obj
            End Function

            Private Shared Iterator Function Enumerate(reference As CalcReference) As IEnumerable(Of Object)
                For Each range In reference.GetRanges()
                    Dim rowCount As Integer = range.Rows.Count
                    Dim colCount As Integer = range.Columns.Count
                    For i As Integer = 0 To rowCount - 1
                        For j As Integer = 0 To colCount - 1
                            If range.Cells(i, j).DisplayFormat.Interior.Color = System.Drawing.Color.Red Then
                                Yield range.Cells(i, j).Value
                            End If
                        Next j
                    Next i
                Next range
            End Function
        End Class

    End Class
End Namespace
