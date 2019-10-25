Imports System.Drawing

Namespace Features.RangeOperations
    Public Class CopyPasteOptions
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set data of PC
            worksheet.Range!A2.Value = "PC"
            worksheet.Range("A4:C4").Value = New String() {"Device", "Quantity", "Unit Price"}
            worksheet.Range("A5:C10").Value = New Object(,) {
                {"T540p", 12, 9850},
                {"T570", 5, 7460},
                {"Y460", 6, 5400},
                {"Y460F", 8, 6240}
            }

            'Set style
            worksheet.Range!A2.RowHeight = 30
            worksheet.Range!A2.Font.Size = 20
            worksheet.Range!A2.Font.Bold = True
            worksheet.Range("A4:C4").Font.Bold = True
            worksheet.Range("A4:C4").Font.Color = Color.White
            worksheet.Range("A4:C4").Interior.Color = Color.LightBlue
            worksheet.Range("A5:C10").Borders(BordersIndex.InsideHorizontal).Color = Color.Orange
            worksheet.Range("A5:C10").Borders(BordersIndex.InsideHorizontal).LineStyle = BorderLineStyle.DashDot

            'Copy only style and row height
            worksheet.Range!H1.Value = "Copy style and row height from previous cells."
            worksheet.Range!H1.Font.Color = Color.Red
            worksheet.Range!H1.Font.Bold = True
            worksheet.Range("A2:C10").Copy(worksheet.Range!H2, PasteType.Formats)

            'Set data of mobile devices
            worksheet.Range!H2.Value = "Mobile"
            worksheet.Range("H4:J4").Value = New String() {"Device", "Quantity", "Unit Price"}
            worksheet.Range("H5:J10").Value = New Object(,) {
                {"HW-P30", 20, 4200},
                {"IPhone-X", 5, 9888},
                {"IPhone-6s plus", 15, 6880}
            }

            'Add new sheet
            Dim worksheet2 As IWorksheet = workbook.Worksheets.Add()

            'Copy only style to new sheet
            worksheet.Range("A2:C10").Copy(worksheet2.Range!A2, PasteType.Formats)
            worksheet2.Range!A3.Value = "Copy style from sheet1."
            worksheet2.Range!A3.Font.Color = Color.Red
            worksheet2.Range!A3.Font.Bold = True
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
