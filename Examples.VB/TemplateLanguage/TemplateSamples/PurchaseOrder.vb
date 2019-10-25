Imports System.Data

Namespace Templates.TemplateSamples
    Public Class PurchaseOrder
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_PurchaseOrder.xlsx")
            workbook.Open(templateFile)

#Region "Define custom classes"
            'Public Class PurchaseOrderBasicInfo
            '    Public ID As String
            '    Public OrderDate As Date
            '    Public CreditTerms As String
            '    Public PONumber As String
            '    Public Ref As String
            '    Public DeliverToCompany As String
            '    Public DeliverToAddress As String
            '    Public PostalCode As String
            '    Public Country As String
            'End Class
#End Region

#Region "Init Data"
            Dim po As New DataTable

            With po.Columns
                .Add(New DataColumn("s_no", GetType(Integer)))
                .Add(New DataColumn("itemnumber", GetType(String)))
                .Add(New DataColumn("itemdescription", GetType(String)))
                .Add(New DataColumn("quantity", GetType(Integer)))
                .Add(New DataColumn("um", GetType(String)))
                .Add(New DataColumn("price", GetType(Integer)))
            End With

            With po.Rows
                .Add(1, "P1001", "Pencils HB", 5, "dozen", 10)
                .Add(2, "P1003", "Pencils 2B", 4, "dozen", 10)
                .Add(3, "P1003", "Paper A4 - Photo Copier", 10, "ream", 3)
                .Add(4, "P1234", "Pens - Ball point", 15, "boxes", 2)
                .Add(5, "P3221", "Highligter", 8, "sets", 10)
            End With

            Dim orderbasicInfo As New PurchaseOrderBasicInfo With {
                .ID = "US120499",
                .OrderDate = #2019-7-7#,
                .CreditTerms = "30",
                .PONumber = "PO1011",
                .Ref = "QT1231",
                .DeliverToCompany = "Sanfort Pvt. Ltd.",
                .DeliverToAddress = "1322, High Street, Geln Waverlay",
                .PostalCode = "Victoria 3456",
                .Country = "Australia"
            }
#End Region

            'Add data source
            workbook.AddDataSource("po", po)
            workbook.AddDataSource("tax", 5)
            workbook.AddDataSource("ds", orderbasicInfo)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_PurchaseOrder.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property HasTemplate As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanDownloadZip As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Template_PurchaseOrder.xlsx"}
            End Get
        End Property
    End Class

    Public Class PurchaseOrderBasicInfo
        Public ID As String
        Public OrderDate As Date
        Public CreditTerms As String
        Public PONumber As String
        Public Ref As String
        Public DeliverToCompany As String
        Public DeliverToAddress As String
        Public PostalCode As String
        Public Country As String
    End Class
End Namespace
