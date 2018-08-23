Friend Class ExampleComparer
    Implements IComparer(Of ExampleBase)
    Private _sortOrders As New Dictionary(Of String, String)()
    Public Sub New()
        ' root children orders
        _sortOrders.Add("Tutorial", "a")
        _sortOrders.Add("Features", "b")
        _sortOrders.Add("SpreadSheetsViewer", "c")
        _sortOrders.Add("ExcelReporting", "d")
        _sortOrders.Add("ExcelTemplates", "e")
        ' Features children orders
        _sortOrders.Add("RangeOperations", "a")
        _sortOrders.Add("Formatting", "b")
        _sortOrders.Add("Tables", "c")
        _sortOrders.Add("ConditionalFormatting", "d")
        _sortOrders.Add("DataValidation", "e")
        _sortOrders.Add("Formulas", "f")
        _sortOrders.Add("Grouping", "g")
        _sortOrders.Add("Filtering", "h")
        _sortOrders.Add("Sorting", "i")
        _sortOrders.Add("Sparklines", "j")
        _sortOrders.Add("Charts", "k")
        _sortOrders.Add("Shape", "l")
        _sortOrders.Add("Picture", "m")
        _sortOrders.Add("Slicer", "n")
        _sortOrders.Add("Comments", "o")
        _sortOrders.Add("PivotTable", "p")
        _sortOrders.Add("Hyperlinks", "q")
        _sortOrders.Add("Theme", "r")
        _sortOrders.Add("Workbook", "s")
        _sortOrders.Add("Worksheets", "t")
    End Sub
    Public Function Compare(x As ExampleBase, y As ExampleBase) As Integer Implements IComparer(Of ExampleBase).Compare
        If TypeOf x Is Tutorial Then
            Return -1
        ElseIf TypeOf y Is Tutorial Then
            Return 1
        End If
        Dim xSortKey As String = Nothing
        If Not _sortOrders.TryGetValue(x.GetShortID, xSortKey) Then
            xSortKey = x.ID
        End If
        Dim ySortKey As String = Nothing
        If Not _sortOrders.TryGetValue(y.GetShortID, ySortKey) Then
            ySortKey = y.ID
        End If
        If TypeOf x Is FolderExample Then
            If TypeOf y Is FolderExample Then
                Return xSortKey.CompareTo(ySortKey)
            Else
                Return -1
            End If
        Else
            If TypeOf y Is FolderExample Then
                Return 1
            Else
                Return xSortKey.CompareTo(ySortKey)
            End If
        End If
    End Function
End Class
