Imports System.Data

Namespace Templates.TemplateSamples
    Public Class Travel
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_Travel.xlsx")
            workbook.Open(templateFile)

#Region "Init Data"
            Dim ds1 As New DataTable
            With ds1.Columns
                .Add(New DataColumn("Carrier", GetType(String)))
                .Add(New DataColumn("FlightNo", GetType(Integer)))
                .Add(New DataColumn("Date", GetType(Date)))
                .Add(New DataColumn("From", GetType(String)))
                .Add(New DataColumn("DepartureTime", GetType(TimeSpan)))
                .Add(New DataColumn("To", GetType(String)))
                .Add(New DataColumn("ArrivalTime", GetType(TimeSpan)))
                .Add(New DataColumn("ReservationNo", GetType(String)))
            End With

            ds1.Rows.Add("Trenz Airlines", 1623, #10/25/2018#, "Lorem International",
                         New TimeSpan(7, 56, 0), "Dolor Airport",
                         New TimeSpan(9, 15, 0), "AG4567997")
            ds1.Rows.Add("Trenz Airlines", 1323, #10/30/2018#, "Lorem International",
                         New TimeSpan(20, 25, 0), "Dolor Airport",
                         New TimeSpan(21, 45, 0), "AG4567998")

            Dim ds2 As New DataTable
            With ds2.Columns
                .Add(New DataColumn("Accommodations", GetType(String)))
                .Add(New DataColumn("Date", GetType(Date)))
                .Add(New DataColumn("Concierge", GetType(String)))
                .Add(New DataColumn("Phone", GetType(String)))
                .Add(New DataColumn("Email", GetType(String)))
                .Add(New DataColumn("AddressPart1", GetType(String)))
                .Add(New DataColumn("AddressPart2", GetType(String)))
                .Add(New DataColumn("ConfirmNo", GetType(String)))
                .Add(New DataColumn("Days", GetType(Integer)))
                .Add(New DataColumn("TotalCost", GetType(Double)))
            End With

            ds2.Rows.Add("Lorem Hotel", #10/25/2018#, "Charles", "01234 567 890",
                         "charles@lorem.com", "123 High Street, ",
                         "Anytown, County, Postcode", "A4567", 2, 800)
            ds2.Rows.Add("Deloz Hotel", #10/27/2018#, "James",
                         "01234 567 890", "no_reply@example.com",
                         "202 Halford Street, ", "Anytown, County, Postcode", "A4568", 3, 900)

            Dim ds3 As New DataTable
            ds3.Columns.Add(New DataColumn("Contact", GetType(String)))
            ds3.Columns.Add(New DataColumn("Phone", GetType(String)))

            ds3.Rows.Add("Airline Reservations", "01234 567 890")
            ds3.Rows.Add("Hotel Reservations", "12342322232")

            Dim ds4 As New DataTable
            With ds4.Columns
                .Add(New DataColumn("Contact", GetType(String)))
                .Add(New DataColumn("Phone", GetType(String)))
                .Add(New DataColumn("Notes", GetType(String)))
            End With

            ds4.Rows.Add("Tom Jenkins", "01234 567 890", "tom.jerkins@trenz.com")
            ds4.Rows.Add("Rayna James", "19234222456", "ratna.james@deloz.com")
#End Region

            'Add data source
            workbook.AddDataSource("ds1", ds1)
            workbook.AddDataSource("ds2", ds2)
            workbook.AddDataSource("ds3", ds3)
            workbook.AddDataSource("ds4", ds4)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_Travel.xlsx"
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
                Return New String() {"xlsx\Template_Travel.xlsx"}
            End Get
        End Property
    End Class


End Namespace
