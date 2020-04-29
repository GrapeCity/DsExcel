Namespace ExcelTemplates
    Public Class EventBudget
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Event budget.xlsx from resource
            Dim fileStream = GetResourceStream("Event budget.xlsx")

            workbook.Open(fileStream)

            Dim worksheet = workbook.ActiveSheet

            'change range B2's font size.
            worksheet.Range!B2.Font.Size = 22

            'change range E4's font style to bold.
            worksheet.Range!E4.Font.Bold = True

            'change table style.
            worksheet.Tables("tblAdmissions").TableStyle = workbook.TableStyles("TableStyleLight10")
            worksheet.Tables("tblAds").TableStyle = workbook.TableStyles("TableStyleLight10")
            worksheet.Tables("tblVendors").TableStyle = workbook.TableStyles("TableStyleLight10")
            worksheet.Tables("tblItems").TableStyle = workbook.TableStyles("TableStyleLight10")

            'modify range F4:G5's cell style.
            worksheet.Range("F4:G5").Interior.ThemeColor = ThemeColor.Light1
            worksheet.Range("F4:G5").Interior.TintAndShade = -0.15
            worksheet.Range("F4:G5").Font.ThemeFont = ThemeFont.Major
            worksheet.Range("F4:G5").Font.Size = 12
            worksheet.Range("F4:G5").Borders(BordersIndex.InsideHorizontal).LineStyle = BorderLineStyle.None
            worksheet.Range("F5:G5").NumberFormat = "$#,##0.00"

            'modify table columns' style.
            worksheet.Range("F8:G11, F15:G18, F22:G25, F29:G33").Interior.ThemeColor = ThemeColor.Light1
            worksheet.Range("F8:G11, F15:G18, F22:G25, F29:G33").Interior.TintAndShade = -0.15
            worksheet.Range("E8:G11, E15:G18, E22:G25, E29:G33").NumberFormat = "$#,##0.00"
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Event budget.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property HasTemplate As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
