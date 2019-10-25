Imports System.Drawing

Namespace Features.Formulas.Precedentsanddependents
    Public Class GetAllPrecedents
        Inherits ExampleBase

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            With worksheet.Range
                !E2.Formula = "=sum(C1:C2)"
                !C1.Formula = "=B1"
                !B1.Formula = "=sum(A1:A2)"
                !A1.Value = 1
                !A2.Value = 2
                !C2.Value = 3
            End With

            ' Add precedents of E2
            Dim precedentsList = worksheet.Range!E2.GetPrecedents.ToList

            ' Add inner precedents of E2
            Dim nextLayer = precedentsList
            Do While nextLayer.Count > 0
                Dim currentLayer = nextLayer
                nextLayer = New List(Of IRange)

                For Each precedentRange In currentLayer
                    For i = 0 To precedentRange.RowCount - 1
                        For j = 0 To precedentRange.ColumnCount - 1
                            Dim innerPrecedents = precedentRange.Cells(i, j).GetPrecedents()
                            If innerPrecedents.Count = 0 Then
                                precedentRange.Cells(i, j).Interior.Color = Color.SkyBlue
                            Else
                                precedentRange.Cells(i, j).Interior.Color = Color.Gray
                                nextLayer.AddRange(innerPrecedents)
                            End If
                        Next j
                    Next i
                Next precedentRange
            Loop
        End Sub

    End Class
End Namespace
