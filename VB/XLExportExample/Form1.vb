Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet
Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Namespace XLExportExample
    Partial Public Class Form1
        Inherits Form

        Private invoice As Invoice
        Private panelFont As XlFont
        Private titleFont As XlFont
        Private infoFont As XlFont
        Private leftPanelBorder As XlBorder
        Private leftPanelFormatting As XlCellFormatting
        Private rightPanelFormatting As XlCellFormatting
        Private headerRowFormatting As XlCellFormatting
        Private evenRowFormatting As XlCellFormatting
        Private oddRowFormatting As XlCellFormatting
        Private infoFormatting As XlCellFormatting

        Public Sub New()
            InitializeComponent()
            InitializeFormatting()
            invoice = XLExportExample.Invoice.CreateSampleInvoice()
        End Sub

        Private Sub InitializeFormatting()
            ' Specify formatting settings for the even rows.
            evenRowFormatting = New XlCellFormatting()
            evenRowFormatting.Font = New XlFont()
            evenRowFormatting.Font.Name = "Century Gothic"
            evenRowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
            evenRowFormatting.Fill = XlFill.SolidFill(Color.White)

            ' Specify formatting settings for the odd rows.
            oddRowFormatting = New XlCellFormatting()
            oddRowFormatting.CopyFrom(evenRowFormatting)
            oddRowFormatting.Fill = XlFill.SolidFill(Color.FromArgb(242, 242, 242))

            ' Specify formatting settings for the header row.
            headerRowFormatting = New XlCellFormatting()
            headerRowFormatting.CopyFrom(evenRowFormatting)
            headerRowFormatting.Font.Bold = True
            headerRowFormatting.Font.Color = Color.White
            headerRowFormatting.Fill = XlFill.SolidFill(Color.FromArgb(192, 0, 0))
            ' Set borders for the header row.
            headerRowFormatting.Border = New XlBorder()
            ' Specify the top border and set its color to white.
            headerRowFormatting.Border.TopColor = Color.White
            ' Specify the medium border line style. 
            headerRowFormatting.Border.TopLineStyle = XlBorderLineStyle.Medium
            ' Specify the bottom border for the header row.
            ' Set the bottom border color to dark gray.
            headerRowFormatting.Border.BottomColor = Color.FromArgb(89, 89, 89)
            ' Specify the medium border line style.
            headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Medium

            ' Specify formatting settings for the invoice header.
            panelFont = New XlFont()
            panelFont.Name = "Century Gothic"
            panelFont.SchemeStyle = XlFontSchemeStyles.None
            panelFont.Color = Color.White

            ' Set font attributes for the row displaying the invoice label and company name. 
            titleFont = panelFont.Clone()
            titleFont.Size = 26

            ' Specify formatting settings for the worksheet range containing the name and contact details of the seller (the "Vader Enterprises" panel).
            leftPanelFormatting = New XlCellFormatting()
            ' Set the cell background color to dark gray.
            leftPanelFormatting.Fill = XlFill.SolidFill(Color.FromArgb(89, 89, 89))
            leftPanelFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom)
            leftPanelFormatting.NumberFormat = XlNumberFormat.General
            ' Set the right border for this range.
            leftPanelBorder = New XlBorder()
            ' Set the right border color to white.
            leftPanelBorder.RightColor = Color.White
            ' Specify the medium border line style. 
            leftPanelBorder.RightLineStyle = XlBorderLineStyle.Medium

            ' Specify formatting settings for the worksheet range containing general information about the invoice: 
            ' its date, reference number and service description (the "Invoice" panel).
            rightPanelFormatting = New XlCellFormatting()
            ' Set the cell background color to dark red.
            rightPanelFormatting.Fill = XlFill.SolidFill(Color.FromArgb(192, 0, 0))
            rightPanelFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom)
            rightPanelFormatting.NumberFormat = XlNumberFormat.General

            ' Specify formatting settings and font attributes for the worksheet range containing buyer's contact information (the "Bill To" panel). 
            infoFormatting = New XlCellFormatting()
            ' Set the cell background color to light gray.
            infoFormatting.Fill = XlFill.SolidFill(Color.FromArgb(217, 217, 217))
            infoFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom)
            infoFormatting.NumberFormat = XlNumberFormat.General
            infoFont = panelFont.Clone()
            infoFont.Color = Color.Black
        End Sub

        ' Export the document to XLSX format.
        Private Sub btnExportToXLSX_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLSX.Click
            Dim fileName As String = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            If ExportToFile(fileName, XlDocumentFormat.Xlsx) Then
                ShowFile(fileName)
            End If
        End Sub

        ' Export the document to XLS format.
        Private Sub btnExportToXLS_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLS.Click
            Dim fileName As String = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            If ExportToFile(fileName, XlDocumentFormat.Xls) Then
                ShowFile(fileName)
            End If
        End Sub

        Private Function GetSaveFileName(ByVal filter As String, ByVal defaulName As String) As String
            saveFileDialog1.Filter = filter
            saveFileDialog1.FileName = defaulName
            If saveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                Return Nothing
            End If
            Return saveFileDialog1.FileName
        End Function

        Private Sub ShowFile(ByVal fileName As String)
            If Not File.Exists(fileName) Then
                Return
            End If
            Dim dResult As DialogResult = MessageBox.Show(String.Format("Do you want to open the resulting file?", fileName), Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If dResult = System.Windows.Forms.DialogResult.Yes Then
                Process.Start(fileName)
            End If
        End Sub

        Private Function ExportToFile(ByVal fileName As String, ByVal documentFormat As XlDocumentFormat) As Boolean
            Try
                Using stream As New FileStream(fileName, FileMode.Create)
                    ' Create an exporter with the specified formula parser.
                    Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat, New XlFormulaParser())
                    ' Create a new document and begin to write it to the specified stream.
                    Using document As IXlDocument = exporter.CreateDocument(stream)
                        ' Generate the document content. 
                        GenerateDocument(document)
                    End Using
                End Using
                Return True
            Catch ex As Exception
                MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        Private Sub GenerateDocument(ByVal document As IXlDocument)
            ' Specify the document culture.
            document.Options.Culture = CultureInfo.CurrentCulture

            ' Add a new worksheet to the document.
            Using sheet As IXlSheet = document.CreateSheet()
                ' Specify the worksheet name.
                sheet.Name = "Invoice"

                ' Specify page settings.
                sheet.PageSetup = New XlPageSetup()
                ' Scale the print area to fit to one page wide.
                sheet.PageSetup.FitToPage = True
                sheet.PageSetup.FitToWidth = 1
                sheet.PageSetup.FitToHeight = 0

                ' Generate worksheet columns.
                GenerateColumns(sheet)

                ' Generate data rows containing the invoice heading.
                GenerateInvoiceTitle(sheet)

                ' Generate data rows containing the recipient's contact information.
                GenerateInvoiceBillTo(sheet)

                ' Generate the header row for the table of purchased products.
                GenerateHeaderRow(sheet)

                Dim firstDataRowIndex As Integer = sheet.CurrentRowIndex

                ' Generate the data row for each product in the invoice list and provide the product information: its description, quantity, unit price and so on.
                For i As Integer = 0 To invoice.Items.Count - 1
                    GenerateDataRow(sheet, invoice.Items(i), (i + 1) = invoice.Items.Count)
                Next i

                ' Generate the total row.
                GenerateTotalRow(sheet, firstDataRowIndex)

                ' Generate data rows containing additional information.
                GenerateInfoRow(sheet, "Make all checks payable to Vader Enterprises")
                GenerateInfoRow(sheet, "If you have any questions concerning this invoice, contact Darth Vader, (111)111-1111, darth@vader.com")

                ' Specify the data range to be printed.
                sheet.PrintArea = sheet.DataRange
            End Using
        End Sub

        #Region "Columns"
        Private Sub GenerateColumns(ByVal sheet As IXlSheet)
            Dim currencyFormat As XlNumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
            Dim discountFormat As XlNumberFormat = "0.00%;[Red]-0.00%;;@"

            ' Create the column "A" and set its width. 
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 21
            End Using
            ' Create the column "B" and set its width. 
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 21
            End Using

            ' Create the column "C" containing the "Description" label in the header row and adjust its width. 
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 120
                ' Set the horizontal and vertical alignment of cell content. 
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center))
            End Using

            ' Create the column "D" and adjust its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 263
                ' Set the horizontal and vertical alignment of cell content. 
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center))
            End Using

            ' Create the column "E" containing the "QTY" label in the header row and adjust its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 102
                ' Set the horizontal and vertical alignment of cell content. 
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
            End Using

            ' Create the column "F" containing the "Unit Price" label in the header row and adjust its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 150
                ' Set the horizontal and vertical alignment of cell content. 
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Center))
                ' Apply the currency number format to the column.
                column.ApplyFormatting(currencyFormat)
            End Using

            ' Create the column "G" containing the "Discount" label in the header row and adjust its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 134
                ' Set the horizontal and vertical alignment of cell content.
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
                ' Apply the custom number format to the column.
                column.ApplyFormatting(discountFormat)
            End Using

            ' Create the column "H" containing the "Amount" label in the header row and adjust its width.
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 174
                ' Set the horizontal and vertical alignment of cell content.
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Center))
                ' Apply the currency number format to the column.
                column.ApplyFormatting(currencyFormat)
            End Using

            ' Create the column "I" and set its width. 
            Using column As IXlColumn = sheet.CreateColumn()
                column.WidthInPixels = 21
            End Using
        End Sub
        #End Region

        #Region "Invoice Heading"
        Private Sub GenerateInvoiceTitle(ByVal sheet As IXlSheet)
            ' Create the empty row at the top of the worksheet.
            Using row As IXlRow = sheet.CreateRow()
            End Using
            ' Create the row containing the company name and invoice label. 
            ' Set the row height to 58 pixels and specify font attributes of cell content.  
            GenerateTitleRow(sheet, "Vader Enterprises", Nothing, "INVOICE", 58, titleFont, Nothing)
            ' Create the empty row with the default height and specific formatting.
            GenerateTitleRow(sheet, Nothing, Nothing, Nothing, -1, panelFont, Nothing)
            ' Create the row containing the company address and invoice date. 
            ' Set the default row height, specify font attributes and number format of cell content.  
            GenerateTitleRow(sheet, "123 Home Lane", "DATE", invoice.Date, -1, panelFont, "mmmm d, yyyy")
            ' Create the row containing the company location and invoice number. 
            ' Set the default row height and specify font attributes of cell content.  
            GenerateTitleRow(sheet, "Homesville, CA, 55555", "INVOICE#", invoice.InvoiceNum, -1, panelFont, Nothing)
            ' Create the row containing the company phone number and service description. 
            ' Set the default row height and specify font attributes of cell content.
            GenerateTitleRow(sheet, "Phone: (111)111-1111, Fax: (111)111-1112", "FOR", "Service description", -1, panelFont, Nothing)
            ' Create the empty row with the default height and specific formatting.
            GenerateTitleRow(sheet, Nothing, Nothing, Nothing, -1, panelFont, Nothing)
        End Sub

        Private Sub GenerateTitleRow(ByVal sheet As IXlSheet, ByVal info As String, ByVal name As String, ByVal value As Object, ByVal rowHeight As Integer, ByVal font As XlFont, ByVal specificFormat As XlNumberFormat)
            Using row As IXlRow = sheet.CreateRow()
                ' Set the row height.
                row.HeightInPixels = rowHeight
                ' Set the cell font.
                row.ApplyFormatting(font)

                ' Create the first empty cell.
                row.SkipCells(1)

                ' Create the blank cell with the specified formatting settings.
                row.BlankCells(1, leftPanelFormatting)

                ' Create the third cell, assign its value and apply specific formatting settings to it.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = info
                    cell.ApplyFormatting(leftPanelFormatting)
                End Using

                ' Create two blank cells with the specified formatting settings.
                row.BlankCells(2, leftPanelFormatting)

                ' Create the cell, apply specific formatting settings to it and set the cell right border.
                Using cell As IXlCell = row.CreateCell()
                    cell.ApplyFormatting(leftPanelFormatting)
                    cell.ApplyFormatting(leftPanelBorder)
                End Using

                ' Create the cell, assign its value and apply specific formatting settings to it.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = name
                    cell.ApplyFormatting(rightPanelFormatting)
                    cell.Formatting.Alignment.Indent = 1
                End Using

                ' Create the cell, assign its value converted from the custom object and apply specific formatting settings to it.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = XlVariantValue.FromObject(value)
                    cell.ApplyFormatting(rightPanelFormatting)
                    If specificFormat IsNot Nothing Then
                        cell.ApplyFormatting(specificFormat)
                    End If
                End Using

                ' Create one blank cell with the specified formatting settings.
                row.BlankCells(1, rightPanelFormatting)
            End Using
        End Sub
        #End Region

        #Region "Invoice BillTo"
        Private Sub GenerateInvoiceBillTo(ByVal sheet As IXlSheet)
            ' Set the top border for the first row in the "Bill To" panel.
            Dim border As New XlBorder()
            ' Set the top border color to white.
            border.TopColor = Color.White
            ' Specify the medium border line style.
            border.TopLineStyle = XlBorderLineStyle.Medium

            ' Generate worksheet rows containing buyer's contact information. 
            GenerateBillToRow(sheet, Nothing, Nothing, Nothing, Nothing, border)
            GenerateBillToRow(sheet, "BILL TO:", invoice.Customer, "PHONE:", invoice.Phone, Nothing)
            GenerateBillToRow(sheet, Nothing, invoice.Company, Nothing, Nothing, Nothing)
            GenerateBillToRow(sheet, "ADDRESS:", invoice.Address, "FAX:", invoice.Fax, Nothing)
            GenerateBillToRow(sheet, Nothing, invoice.Address2, Nothing, Nothing, Nothing)
            GenerateBillToRow(sheet, Nothing, Nothing, Nothing, Nothing, Nothing)
        End Sub

        Private Sub GenerateBillToRow(ByVal sheet As IXlSheet, ByVal name1 As String, ByVal value1 As Object, ByVal name2 As String, ByVal value2 As Object, ByVal specificBorder As XlBorder)
            Using row As IXlRow = sheet.CreateRow()
                ' Set the cell font.
                row.ApplyFormatting(infoFont)
                ' Skip the first cell in the row.
                row.SkipCells(1)

                ' Create the empty cell with the specified formatting settings.
                Using cell As IXlCell = row.CreateCell()
                    cell.ApplyFormatting(infoFormatting)
                    ' Set the cell border.
                    cell.ApplyFormatting(specificBorder)
                End Using

                ' Create the cell, assign its value and apply specific formatting settings to it.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = name1
                    cell.ApplyFormatting(infoFormatting)
                    ' Set the cell border.
                    cell.ApplyFormatting(specificBorder)
                    cell.Formatting.Font.Bold = True
                End Using

                ' Create the cell, assign its value converted from the custom object and apply specific formatting settings to it.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = XlVariantValue.FromObject(value1)
                    cell.ApplyFormatting(infoFormatting)
                    ' Set the cell border.
                    cell.ApplyFormatting(specificBorder)
                End Using

                ' Create the cell, assign its value and apply specific formatting settings to it.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = name2
                    cell.ApplyFormatting(infoFormatting)
                    ' Set the cell border.
                    cell.ApplyFormatting(specificBorder)
                    cell.Formatting.Font.Bold = True
                End Using

                ' Create the cell, assign its value converted from the custom object and apply specific formatting settings to it.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = XlVariantValue.FromObject(value2)
                    cell.ApplyFormatting(infoFormatting)
                    ' Set the cell border.
                    cell.ApplyFormatting(specificBorder)
                End Using

                ' Create three successive cells, apply specific formatting settings to them and set the cell borders.
                For i As Integer = 0 To 2
                    Using cell As IXlCell = row.CreateCell()
                        cell.ApplyFormatting(infoFormatting)
                        cell.ApplyFormatting(specificBorder)
                    End Using
                Next i
            End Using
        End Sub
        #End Region

        #Region "Invoice Content"
        Private Sub GenerateHeaderRow(ByVal sheet As IXlSheet)
            ' Create an array that contains column labels for the header row.
            Dim columnNames() As String = { "Description", Nothing, "QTY", "Unit Price", "Discount", "Amount" }
            ' Create the header row.
            Using row As IXlRow = sheet.CreateRow()
                ' Set the row height to 28 pixels. 
                row.HeightInPixels = 28
                ' Skip the first cell in the row.
                row.SkipCells(1)
                ' Create one blank cell with the specified formatting settings.
                row.BlankCells(1, headerRowFormatting)

                ' Create cells that display column labels and apply specific formatting settings to them.
                For Each columnName As String In columnNames
                    Using cell As IXlCell = row.CreateCell()
                        cell.Value = columnName
                        cell.ApplyFormatting(headerRowFormatting)
                    End Using
                Next columnName

                ' Create one blank cell with the specified formatting settings.
                row.BlankCells(1, headerRowFormatting)
            End Using
        End Sub

        Private Sub GenerateDataRow(ByVal sheet As IXlSheet, ByVal item As InvoiceData, ByVal isLastRow As Boolean)
            ' Create the data row to display the invoice information on each product.
            Using row As IXlRow = sheet.CreateRow()
                ' Set the row height to 28 pixels.
                row.HeightInPixels = 28

                ' Specify formatting settings to be applied to the data rows to shade alternate rows.
                Dim formatting As New XlCellFormatting()
                formatting.CopyFrom(If(row.RowIndex Mod 2 = 0, evenRowFormatting, oddRowFormatting))
                ' Set the bottom border for the last data row.
                If isLastRow Then
                    formatting.Border = New XlBorder()
                    formatting.Border.BottomColor = Color.FromArgb(89, 89, 89)
                    formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium
                End If

                ' Skip the first cell in the row.
                row.SkipCells(1)
                ' Create the blank cell with the specified formatting settings.
                row.BlankCells(1, formatting)

                ' Create the cell containing the product description. 
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = item.Product
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the blank cell with the specified formatting settings.
                row.BlankCells(1, formatting)

                ' Create the cell containing the product quantity. 
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = item.Qty
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing the unit price.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = item.UnitPrice
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing the product discount.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = item.Discount
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the cell containing the amount.
                Using cell As IXlCell = row.CreateCell()
                    ' Set the cell value.
                    cell.Value = item.Qty * item.UnitPrice * (1 - item.Discount)
                    ' Set the formula to calculate the amount per product.
                    cell.SetFormula(String.Format("E{0}*F{0}*(1-G{0})", cell.RowIndex + 1))
                    cell.ApplyFormatting(formatting)
                End Using

                ' Create the blank cell with the specified formatting settings.
                row.BlankCells(1, formatting)
            End Using
        End Sub

        Private Sub GenerateTotalRow(ByVal sheet As IXlSheet, ByVal firstDataRowIndex As Integer)
            ' Skip one row before starting to generate the total row.
            sheet.SkipRows(1)

            ' Create the total row.
            Using row As IXlRow = sheet.CreateRow()
                ' Set the row height to 28 pixels.
                row.HeightInPixels = 28
                ' Set font characteristics for the row cells.
                row.ApplyFormatting(infoFont.Clone())
                row.Formatting.Font.Bold = True

                ' Skip six successive cells in the total row.
                row.SkipCells(6)

                ' Create the "Total" cell.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = "TOTAL"
                End Using

                ' Create the cell that displays the total amount.
                Using cell As IXlCell = row.CreateCell()
                    ' Set the formula to calculate the total amount.
                    cell.SetFormula(String.Format("SUM(H{0}:H{1})", firstDataRowIndex + 1, row.RowIndex - 1))
                    ' Set the cell background color.
                    cell.ApplyFormatting(XlFill.SolidFill(Color.FromArgb(217, 217, 217)))
                End Using

                ' Create the empty cell.
                Using cell As IXlCell = row.CreateCell()
                    ' Set the cell background color.
                    cell.ApplyFormatting(XlFill.SolidFill(Color.FromArgb(217, 217, 217)))
                End Using
            End Using
        End Sub

        Private Sub GenerateInfoRow(ByVal sheet As IXlSheet, ByVal info As String)
            ' Skip one row before starting to generate the row with additional information.
            sheet.SkipRows(1)

            ' Create the row.
            Using row As IXlRow = sheet.CreateRow()
                ' Skip the first cell in the row.
                row.SkipCells(1)

                ' Create the cell that contains the invoice payment options and set its font attributes.
                Using cell As IXlCell = row.CreateCell()
                    cell.Value = info
                    cell.ApplyFormatting(infoFont)
                End Using
            End Using
        End Sub
        #End Region
    End Class
End Namespace
