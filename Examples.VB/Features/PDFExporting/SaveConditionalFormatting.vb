Imports System.Drawing

Namespace Features.PDFExporting
    Public Class SaveConditionalFormatting
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            'Conditional formatting on merge cell
            sheet.Range("B2:C4").Merge()
            sheet.Range("B2:C4").Value = 123

            Dim cf = DirectCast(sheet.Range("B2:C4").FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Greater, 0), IFormatCondition)
            cf.Borders.ThemeColor = ThemeColor.Accent1
            cf.Borders.LineStyle = BorderLineStyle.Thin

            'Set cell values
            Dim data = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
            sheet.Range("B10:B19").Value = data
            sheet.Range("C10:C19").Value = data
            sheet.Range("D10:D19").Value = data

            'Set conditional formattings

            'Color scale
            Dim cf1 As IColorScale = sheet.Range("B10:B19").FormatConditions.AddColorScale(ColorScaleType.ThreeColorScale)
            cf1.ColorScaleCriteria(0).Type = ConditionValueTypes.LowestValue
            cf1.ColorScaleCriteria(0).FormatColor.Color = Color.FromArgb(248, 105, 107)
            cf1.ColorScaleCriteria(1).Type = ConditionValueTypes.Percentile
            cf1.ColorScaleCriteria(1).Value = 50
            cf1.ColorScaleCriteria(1).FormatColor.Color = Color.FromArgb(255, 235, 132)
            cf1.ColorScaleCriteria(2).Type = ConditionValueTypes.HighestValue
            cf1.ColorScaleCriteria(2).FormatColor.Color = Color.FromArgb(99, 190, 123)

            'Data bar
            sheet.Range("C14").Value = -5
            sheet.Range("C17").Value = -8

            Dim cf2 As IDataBar = sheet.Range("C10:C19").FormatConditions.AddDatabar()
            cf2.MinPoint.Type = ConditionValueTypes.AutomaticMin
            cf2.MaxPoint.Type = ConditionValueTypes.AutomaticMax
            cf2.BarFillType = DataBarFillType.Gradient
            cf2.BarColor.Color = Color.FromArgb(0, 138, 239)
            cf2.BarBorder.Color.Color = Color.FromArgb(0, 138, 239)
            cf2.NegativeBarFormat.Color.Color = Color.FromArgb(255, 0, 0)
            cf2.NegativeBarFormat.BorderColorType = DataBarNegativeColorType.Color
            cf2.NegativeBarFormat.BorderColor.Color = Color.FromArgb(255, 0, 0)
            cf2.AxisColor.Color = Color.Black
            cf2.AxisPosition = DataBarAxisPosition.Automatic

            'Icon set
            Dim cf3 As IIconSetCondition = sheet.Range("D10:D19").FormatConditions.AddIconSetCondition()
            cf3.IconSet = workbook.IconSets(IconSetType.Icon3Symbols)
        End Sub
        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
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
    End Class
End Namespace
