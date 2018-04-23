using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace XLExportExample {
    public partial class Form1 : Form {
        Invoice invoice;
        XlFont panelFont;
        XlFont titleFont;
        XlFont infoFont;
        XlBorder leftPanelBorder;
        XlCellFormatting leftPanelFormatting;
        XlCellFormatting rightPanelFormatting;
        XlCellFormatting headerRowFormatting;
        XlCellFormatting evenRowFormatting;
        XlCellFormatting oddRowFormatting;
        XlCellFormatting infoFormatting;

        public Form1() {
            InitializeComponent();
            InitializeFormatting();
            invoice = Invoice.CreateSampleInvoice();
        }

        void InitializeFormatting() {
            // Specify formatting settings for the even rows.
            evenRowFormatting = new XlCellFormatting();
            evenRowFormatting.Font = new XlFont();
            evenRowFormatting.Font.Name = "Century Gothic";
            evenRowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;
            evenRowFormatting.Fill = XlFill.SolidFill(Color.White);

            // Specify formatting settings for the odd rows.
            oddRowFormatting = new XlCellFormatting();
            oddRowFormatting.CopyFrom(evenRowFormatting);
            oddRowFormatting.Fill = XlFill.SolidFill(Color.FromArgb(242, 242, 242));

            // Specify formatting settings for the header row.
            headerRowFormatting = new XlCellFormatting();
            headerRowFormatting.CopyFrom(evenRowFormatting);
            headerRowFormatting.Font.Bold = true;
            headerRowFormatting.Font.Color = Color.White;
            headerRowFormatting.Fill = XlFill.SolidFill(Color.FromArgb(192, 0, 0));
            // Set borders for the header row.
            headerRowFormatting.Border = new XlBorder();
            // Specify the top border and set its color to white.
            headerRowFormatting.Border.TopColor = Color.White;
            // Specify the medium border line style. 
            headerRowFormatting.Border.TopLineStyle = XlBorderLineStyle.Medium;
            // Specify the bottom border for the header row.
            // Set the bottom border color to dark gray.
            headerRowFormatting.Border.BottomColor = Color.FromArgb(89, 89, 89);
            // Specify the medium border line style.
            headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;

            // Specify formatting settings for the invoice header.
            panelFont = new XlFont();
            panelFont.Name = "Century Gothic";
            panelFont.SchemeStyle = XlFontSchemeStyles.None;
            panelFont.Color = Color.White;

            // Set font attributes for the row displaying the invoice label and company name. 
            titleFont = panelFont.Clone();
            titleFont.Size = 26;

            // Specify formatting settings for the worksheet range containing the name and contact details of the seller (the "Vader Enterprises" panel).
            leftPanelFormatting = new XlCellFormatting();
            // Set the cell background color to dark gray.
            leftPanelFormatting.Fill = XlFill.SolidFill(Color.FromArgb(89, 89, 89));
            leftPanelFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom);
            leftPanelFormatting.NumberFormat = XlNumberFormat.General;
            // Set the right border for this range.
            leftPanelBorder = new XlBorder();
            // Set the right border color to white.
            leftPanelBorder.RightColor = Color.White;
            // Specify the medium border line style. 
            leftPanelBorder.RightLineStyle = XlBorderLineStyle.Medium;

            // Specify formatting settings for the worksheet range containing general information about the invoice: 
            // its date, reference number and service description (the "Invoice" panel).
            rightPanelFormatting = new XlCellFormatting();
            // Set the cell background color to dark red.
            rightPanelFormatting.Fill = XlFill.SolidFill(Color.FromArgb(192, 0, 0));
            rightPanelFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom);
            rightPanelFormatting.NumberFormat = XlNumberFormat.General;

            // Specify formatting settings and font attributes for the worksheet range containing buyer's contact information (the "Bill To" panel). 
            infoFormatting = new XlCellFormatting();
            // Set the cell background color to light gray.
            infoFormatting.Fill = XlFill.SolidFill(Color.FromArgb(217, 217, 217));
            infoFormatting.Alignment = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom);
            infoFormatting.NumberFormat = XlNumberFormat.General; 
            infoFont = panelFont.Clone();
            infoFont.Color = Color.Black;
        }

        // Export the document to XLSX format.
        void btnExportToXLSX_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Xlsx))
                ShowFile(fileName);
        }

        // Export the document to XLS format.
        void btnExportToXLS_Click(object sender, EventArgs e) {
            string fileName = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls");
            if (string.IsNullOrEmpty(fileName))
                return;
            if (ExportToFile(fileName, XlDocumentFormat.Xls))
                ShowFile(fileName);
        }

        string GetSaveFileName(string filter, string defaulName) {
            saveFileDialog1.Filter = filter;
            saveFileDialog1.FileName = defaulName;
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                return null;
            return saveFileDialog1.FileName;
        }

        void ShowFile(string fileName) {
            if (!File.Exists(fileName))
                return;
            DialogResult dResult = MessageBox.Show(String.Format("Do you want to open the resulting file?", fileName),
                this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dResult == DialogResult.Yes)
                Process.Start(fileName);
        }

        bool ExportToFile(string fileName, XlDocumentFormat documentFormat) {
            try {
                using (FileStream stream = new FileStream(fileName, FileMode.Create)) {
                    // Create an exporter with the specified formula parser.
                    IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());
                    // Create a new document and begin to write it to the specified stream.
                    using (IXlDocument document = exporter.CreateDocument(stream)) {
                        // Generate the document content. 
                        GenerateDocument(document);
                    }
                }
                return true;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        void GenerateDocument(IXlDocument document) {
            // Specify the document culture.
            document.Options.Culture = CultureInfo.CurrentCulture;

            // Add a new worksheet to the document.
            using (IXlSheet sheet = document.CreateSheet()) {
                // Specify the worksheet name.
                sheet.Name = "Invoice";

                // Specify page settings.
                sheet.PageSetup = new XlPageSetup();
                // Scale the print area to fit to one page wide.
                sheet.PageSetup.FitToPage = true;
                sheet.PageSetup.FitToWidth = 1;
                sheet.PageSetup.FitToHeight = 0;

                // Generate worksheet columns.
                GenerateColumns(sheet);

                // Generate data rows containing the invoice heading.
                GenerateInvoiceTitle(sheet);

                // Generate data rows containing the recipient's contact information.
                GenerateInvoiceBillTo(sheet);

                // Generate the header row for the table of purchased products.
                GenerateHeaderRow(sheet);

                int firstDataRowIndex = sheet.CurrentRowIndex;

                // Generate the data row for each product in the invoice list and provide the product information: its description, quantity, unit price and so on.
                for (int i = 0; i < invoice.Items.Count; i++)
                    GenerateDataRow(sheet, invoice.Items[i], (i + 1) == invoice.Items.Count);

                // Generate the total row.
                GenerateTotalRow(sheet, firstDataRowIndex);

                // Generate data rows containing additional information.
                GenerateInfoRow(sheet, "Make all checks payable to Vader Enterprises");
                GenerateInfoRow(sheet, "If you have any questions concerning this invoice, contact Darth Vader, (111)111-1111, darth@vader.com");

                // Specify the data range to be printed.
                sheet.PrintArea = sheet.DataRange;
            }
        }

        #region Columns
        void GenerateColumns(IXlSheet sheet) {
            XlNumberFormat currencyFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
            XlNumberFormat discountFormat = @"0.00%;[Red]-0.00%;;@";

            // Create the column "A" and set its width. 
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 21;
            // Create the column "B" and set its width. 
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 21;

            // Create the column "C" containing the "Description" label in the header row and adjust its width. 
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 120;
                // Set the horizontal and vertical alignment of cell content. 
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center));
            }

            // Create the column "D" and adjust its width.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 263;
                // Set the horizontal and vertical alignment of cell content. 
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center));
            }

            // Create the column "E" containing the "QTY" label in the header row and adjust its width.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 102;
                // Set the horizontal and vertical alignment of cell content. 
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center));
            }

            // Create the column "F" containing the "Unit Price" label in the header row and adjust its width.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 150;
                // Set the horizontal and vertical alignment of cell content. 
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Center));
                // Apply the currency number format to the column.
                column.ApplyFormatting(currencyFormat);
            }

            // Create the column "G" containing the "Discount" label in the header row and adjust its width.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 134;
                // Set the horizontal and vertical alignment of cell content.
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center));
                // Apply the custom number format to the column.
                column.ApplyFormatting(discountFormat);
            }

            // Create the column "H" containing the "Amount" label in the header row and adjust its width.
            using (IXlColumn column = sheet.CreateColumn()) {
                column.WidthInPixels = 174;
                // Set the horizontal and vertical alignment of cell content.
                column.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Center));
                // Apply the currency number format to the column.
                column.ApplyFormatting(currencyFormat);
            }

            // Create the column "I" and set its width. 
            using (IXlColumn column = sheet.CreateColumn())
                column.WidthInPixels = 21;
        }
        #endregion

        #region Invoice Heading
        void GenerateInvoiceTitle(IXlSheet sheet) {
            // Create the empty row at the top of the worksheet.
            using (IXlRow row = sheet.CreateRow()) { }
            // Create the row containing the company name and invoice label. 
            // Set the row height to 58 pixels and specify font attributes of cell content.  
            GenerateTitleRow(sheet, "Vader Enterprises", null, "INVOICE", 58, titleFont, null);
            // Create the empty row with the default height and specific formatting.
            GenerateTitleRow(sheet, null, null, null, -1, panelFont, null);
            // Create the row containing the company address and invoice date. 
            // Set the default row height, specify font attributes and number format of cell content.  
            GenerateTitleRow(sheet, "123 Home Lane", "DATE", invoice.Date, -1, panelFont, "mmmm d, yyyy");
            // Create the row containing the company location and invoice number. 
            // Set the default row height and specify font attributes of cell content.  
            GenerateTitleRow(sheet, "Homesville, CA, 55555", "INVOICE#", invoice.InvoiceNum, -1, panelFont, null);
            // Create the row containing the company phone number and service description. 
            // Set the default row height and specify font attributes of cell content.
            GenerateTitleRow(sheet, "Phone: (111)111-1111, Fax: (111)111-1112", "FOR", "Service description", -1, panelFont, null);
            // Create the empty row with the default height and specific formatting.
            GenerateTitleRow(sheet, null, null, null, -1, panelFont, null);
        }

        void GenerateTitleRow(IXlSheet sheet, string info, string name, object value, int rowHeight, XlFont font, XlNumberFormat specificFormat) {
            using (IXlRow row = sheet.CreateRow()) {
                // Set the row height.
                row.HeightInPixels = rowHeight;
                // Set the cell font.
                row.ApplyFormatting(font);

                // Create the first empty cell.
                row.SkipCells(1);

                // Create the blank cell with the specified formatting settings.
                row.BlankCells(1, leftPanelFormatting);

                // Create the third cell, assign its value and apply specific formatting settings to it.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = info;
                    cell.ApplyFormatting(leftPanelFormatting);
                }

                // Create two blank cells with the specified formatting settings.
                row.BlankCells(2, leftPanelFormatting);

                // Create the cell, apply specific formatting settings to it and set the cell right border.
                using (IXlCell cell = row.CreateCell()) {
                    cell.ApplyFormatting(leftPanelFormatting);
                    cell.ApplyFormatting(leftPanelBorder);
                }

                // Create the cell, assign its value and apply specific formatting settings to it.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = name;
                    cell.ApplyFormatting(rightPanelFormatting);
                    cell.Formatting.Alignment.Indent = 1;
                }

                // Create the cell, assign its value converted from the custom object and apply specific formatting settings to it.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = XlVariantValue.FromObject(value);
                    cell.ApplyFormatting(rightPanelFormatting);
                    if (specificFormat != null)
                        cell.ApplyFormatting(specificFormat);
                }

                // Create one blank cell with the specified formatting settings.
                row.BlankCells(1, rightPanelFormatting);
            }
        }
        #endregion

        #region Invoice BillTo
        void GenerateInvoiceBillTo(IXlSheet sheet) {
            // Set the top border for the first row in the "Bill To" panel.
            XlBorder border = new XlBorder();
            // Set the top border color to white.
            border.TopColor = Color.White;
            // Specify the medium border line style.
            border.TopLineStyle = XlBorderLineStyle.Medium;

            // Generate worksheet rows containing buyer's contact information. 
            GenerateBillToRow(sheet, null, null, null, null, border);
            GenerateBillToRow(sheet, "BILL TO:", invoice.Customer, "PHONE:", invoice.Phone, null);
            GenerateBillToRow(sheet, null, invoice.Company, null, null, null);
            GenerateBillToRow(sheet, "ADDRESS:", invoice.Address, "FAX:", invoice.Fax, null);
            GenerateBillToRow(sheet, null, invoice.Address2, null, null, null);
            GenerateBillToRow(sheet, null, null, null, null, null);
        }

        void GenerateBillToRow(IXlSheet sheet, string name1, object value1, string name2, object value2, XlBorder specificBorder) {
            using (IXlRow row = sheet.CreateRow()) {
                // Set the cell font.
                row.ApplyFormatting(infoFont);
                // Skip the first cell in the row.
                row.SkipCells(1);

                // Create the empty cell with the specified formatting settings.
                using (IXlCell cell = row.CreateCell()) {
                    cell.ApplyFormatting(infoFormatting);
                    // Set the cell border.
                    cell.ApplyFormatting(specificBorder);
                }

                // Create the cell, assign its value and apply specific formatting settings to it.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = name1;
                    cell.ApplyFormatting(infoFormatting);
                    // Set the cell border.
                    cell.ApplyFormatting(specificBorder);
                    cell.Formatting.Font.Bold = true;
                }

                // Create the cell, assign its value converted from the custom object and apply specific formatting settings to it.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = XlVariantValue.FromObject(value1);
                    cell.ApplyFormatting(infoFormatting);
                    // Set the cell border.
                    cell.ApplyFormatting(specificBorder);
                }

                // Create the cell, assign its value and apply specific formatting settings to it.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = name2;
                    cell.ApplyFormatting(infoFormatting);
                    // Set the cell border.
                    cell.ApplyFormatting(specificBorder);
                    cell.Formatting.Font.Bold = true;
                }

                // Create the cell, assign its value converted from the custom object and apply specific formatting settings to it.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = XlVariantValue.FromObject(value2);
                    cell.ApplyFormatting(infoFormatting);
                    // Set the cell border.
                    cell.ApplyFormatting(specificBorder);
                }

                // Create three successive cells, apply specific formatting settings to them and set the cell borders.
                for (int i = 0; i < 3; i++) {
                    using (IXlCell cell = row.CreateCell()) {
                        cell.ApplyFormatting(infoFormatting);
                        cell.ApplyFormatting(specificBorder);
                    }
                }
            }
        }
        #endregion

        #region Invoice Content
        void GenerateHeaderRow(IXlSheet sheet) {
            // Create an array that contains column labels for the header row.
            string[] columnNames = new string[] { "Description", null, "QTY", "Unit Price", "Discount", "Amount" };
            // Create the header row.
            using (IXlRow row = sheet.CreateRow()) {
                // Set the row height to 28 pixels. 
                row.HeightInPixels = 28;
                // Skip the first cell in the row.
                row.SkipCells(1);
                // Create one blank cell with the specified formatting settings.
                row.BlankCells(1, headerRowFormatting);

                // Create cells that display column labels and apply specific formatting settings to them.
                foreach (string columnName in columnNames) {
                    using (IXlCell cell = row.CreateCell()) {
                        cell.Value = columnName;
                        cell.ApplyFormatting(headerRowFormatting);
                    }
                }

                // Create one blank cell with the specified formatting settings.
                row.BlankCells(1, headerRowFormatting);
            }
        }

        void GenerateDataRow(IXlSheet sheet, InvoiceData item, bool isLastRow) {
            // Create the data row to display the invoice information on each product.
            using (IXlRow row = sheet.CreateRow()) {
                // Set the row height to 28 pixels.
                row.HeightInPixels = 28;

                // Specify formatting settings to be applied to the data rows to shade alternate rows.
                XlCellFormatting formatting = new XlCellFormatting();
                formatting.CopyFrom((row.RowIndex % 2 == 0) ? evenRowFormatting : oddRowFormatting);
                // Set the bottom border for the last data row.
                if (isLastRow) {
                    formatting.Border = new XlBorder();
                    formatting.Border.BottomColor = Color.FromArgb(89, 89, 89);
                    formatting.Border.BottomLineStyle = XlBorderLineStyle.Medium;
                }

                // Skip the first cell in the row.
                row.SkipCells(1);
                // Create the blank cell with the specified formatting settings.
                row.BlankCells(1, formatting);

                // Create the cell containing the product description. 
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = item.Product;
                    cell.ApplyFormatting(formatting);
                }

                // Create the blank cell with the specified formatting settings.
                row.BlankCells(1, formatting);

                // Create the cell containing the product quantity. 
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = item.Qty;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing the unit price.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = item.UnitPrice;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing the product discount.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = item.Discount;
                    cell.ApplyFormatting(formatting);
                }

                // Create the cell containing the amount.
                using (IXlCell cell = row.CreateCell()) {
                    // Set the cell value.
                    cell.Value = item.Qty * item.UnitPrice * (1 - item.Discount);
                    // Set the formula to calculate the amount per product.
                    cell.SetFormula(string.Format("E{0}*F{0}*(1-G{0})", cell.RowIndex + 1));
                    cell.ApplyFormatting(formatting);
                }

                // Create the blank cell with the specified formatting settings.
                row.BlankCells(1, formatting);
            }
        }

        void GenerateTotalRow(IXlSheet sheet, int firstDataRowIndex) {
            // Skip one row before starting to generate the total row.
            sheet.SkipRows(1);

            // Create the total row.
            using (IXlRow row = sheet.CreateRow()) {
                // Set the row height to 28 pixels.
                row.HeightInPixels = 28;
                // Set font characteristics for the row cells.
                row.ApplyFormatting(infoFont.Clone());
                row.Formatting.Font.Bold = true;

                // Skip six successive cells in the total row.
                row.SkipCells(6);

                // Create the "Total" cell.
                using (IXlCell cell = row.CreateCell())
                    cell.Value = "TOTAL";

                // Create the cell that displays the total amount.
                using (IXlCell cell = row.CreateCell()) {
                    // Set the formula to calculate the total amount.
                    cell.SetFormula(string.Format("SUM(H{0}:H{1})", firstDataRowIndex + 1, row.RowIndex - 1));
                    // Set the cell background color.
                    cell.ApplyFormatting(XlFill.SolidFill(Color.FromArgb(217, 217, 217)));
                }

                // Create the empty cell.
                using (IXlCell cell = row.CreateCell())
                    // Set the cell background color.
                    cell.ApplyFormatting(XlFill.SolidFill(Color.FromArgb(217, 217, 217)));
            }
        }

        void GenerateInfoRow(IXlSheet sheet, string info) {
            // Skip one row before starting to generate the row with additional information.
            sheet.SkipRows(1);

            // Create the row.
            using (IXlRow row = sheet.CreateRow()) {
                // Skip the first cell in the row.
                row.SkipCells(1);

                // Create the cell that contains the invoice payment options and set its font attributes.
                using (IXlCell cell = row.CreateCell()) {
                    cell.Value = info;
                    cell.ApplyFormatting(infoFont);
                }
            }
        }
        #endregion
    }
}
