Namespace Templates.TemplateSamples
    Public Class ShippingAndDeliveryOrder
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_ShippingAndDeliveryOrder.xlsx")
            workbook.Open(templateFile)

#Region "Define custom classes"
            'Friend Class PackingList
            '    Public exporter As String
            '    Public address_exporter As String
            '    Public country_exporter As String
            '    Public phonenumber_shipper As String
            '    Public shipper As String

            '    Public [imports] As String
            '    Public address_consignee As String
            '    Public country_consignee As String
            '    Public phonenumber_consignee As String
            '    Public consignee As String

            '    Public invoice_No As Integer
            '    Public [date] As Date
            '    Public reference As Integer

            '    Public dispatchMethod As String
            '    Public shipmentType As String
            '    Public VA As String
            '    Public voyageNo As String
            '    Public portofLanding As String
            '    Public departureDate As Date
            '    Public dischargePort As String
            '    Public finalDestination As String

            '    Public goodsOriginCountry As String
            '    Public destinationCountry As String

            '    Public item As List(Of Product)

            '    Public issuePlace As String
            '    Public issueDate As Date
            '    Public SignatoryCompany As String
            '    Public SignatoryName As String


            'End Class

            'Friend Class Product
            '    Public productcode As String
            '    Public Goods As String
            '    Public quantity As Double
            '    Public netweight As Double
            '    Public kindAndPackagesCount As String
            '    Public grossweight As Double
            '    Public measurements As Double
            'End Class
#End Region

#Region "Init Data"
            Dim packinginfo As New PackingList With {
                .exporter = "DEL Exports",
                .address_exporter = "4243 Longline Vlvd Longline, CA - 98020",
                .country_exporter = "United States",
                .phonenumber_shipper = "010-510-22424",
                .shipper = "Diana Thompson",
                .[imports] = "Deanna Imports",
                .address_consignee = "113/23, Lombard Street Halford Townsville, Melbourne, 4323",
                .country_consignee = "Australia",
                .phonenumber_consignee = "010-510-33232",
                .consignee = "James Williams",
                .invoice_No = 1934,
                .[date] = #2019-1-30#,
                .reference = 1934,
                .dispatchMethod = "Sea",
                .shipmentType = "FCL",
                .goodsOriginCountry = "United States",
                .destinationCountry = "Australia",
                .VA = "MAKERS DYER",
                .voyageNo = "6E",
                .portofLanding = "Longline - California",
                .departureDate = #2019-2-1#,
                .dischargePort = "Melbourne - Australia",
                .finalDestination = "Australia",
                .item = New List(Of Product) From {
                    New Product With {
                        .productcode = "P1001",
                        .Goods = "Pencils - HB",
                        .quantity = 5,
                        .netweight = 0.1,
                        .kindAndPackagesCount = "PALLET X 1",
                        .grossweight = 750,
                        .measurements = 1.7
                    },
                    New Product With {
                        .productcode = "P1002",
                        .Goods = "Paper - A4",
                        .quantity = 3,
                        .netweight = 2,
                        .kindAndPackagesCount = "PALLET X 2",
                        .grossweight = 250,
                        .measurements = 1.4
                    }
                },
                .issuePlace = "Longline",
                .issueDate = #2019-1-30#,
                .SignatoryCompany = "DEL Exports",
                .SignatoryName = "Rayna Johnson"
            }
#End Region

            'Add data source
            workbook.AddDataSource("ds", packinginfo)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_ShippingAndDeliveryOrder.xlsx"
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
                Return New String() {"xlsx\Template_ShippingAndDeliveryOrder.xlsx"}
            End Get
        End Property
    End Class

    Friend Class PackingList
        Public exporter As String
        Public address_exporter As String
        Public country_exporter As String
        Public phonenumber_shipper As String
        Public shipper As String

        Public [imports] As String
        Public address_consignee As String
        Public country_consignee As String
        Public phonenumber_consignee As String
        Public consignee As String

        Public invoice_No As Integer
        Public [date] As Date
        Public reference As Integer

        Public dispatchMethod As String
        Public shipmentType As String
        Public VA As String
        Public voyageNo As String
        Public portofLanding As String
        Public departureDate As Date
        Public dischargePort As String
        Public finalDestination As String

        Public goodsOriginCountry As String
        Public destinationCountry As String

        Public item As List(Of Product)

        Public issuePlace As String
        Public issueDate As Date
        Public SignatoryCompany As String
        Public SignatoryName As String
    End Class

    Friend Class Product
        Public productcode As String
        Public Goods As String
        Public quantity As Double
        Public netweight As Double
        Public kindAndPackagesCount As String
        Public grossweight As Double
        Public measurements As Double
    End Class
End Namespace
