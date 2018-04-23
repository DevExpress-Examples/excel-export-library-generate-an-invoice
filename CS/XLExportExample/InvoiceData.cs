using System;
using System.Collections.Generic;

namespace XLExportExample {
    class InvoiceData 
    {
        public InvoiceData(string product, int qty, double unitPrice, double discount) {
            Product = product;
            Qty = qty;
            UnitPrice = unitPrice;
            Discount = discount;
        }

        public string Product { get; private set; }
        public int Qty { get; private set; }
        public double UnitPrice { get; private set; }
        public double Discount { get; private set; }
    }

    class Invoice 
    {
        readonly List<InvoiceData> items = new List<InvoiceData>();

        public static Invoice CreateSampleInvoice() {
            Invoice result = new Invoice();
            result.InvoiceNum = 100;
            result.Date = DateTime.Now;
            result.Customer = "Alcorn Mickey";
            result.Company = "Mickeys World of Fun";
            result.Address = "436 1st Ave.";
            result.Address2 = "Cleveland, OH, 37288";
            result.Phone = "(203)290-8902";
            result.Fax = "(203)290-8903";
            result.Items.Add(new InvoiceData("Aniseed Syrup", 4, 10, 0));
            result.Items.Add(new InvoiceData("Mishi Kobe Niku", 13, 97, 0.15));
            result.Items.Add(new InvoiceData("Ikura", 12, 31, 0.1));
            result.Items.Add(new InvoiceData("Konbu", 11, 6, 0));
            result.Items.Add(new InvoiceData("Pavlova", 10, 17.45, 0));
            result.Items.Add(new InvoiceData("Boston Crab Meat", 2, 18.4, 0));
            return result;
        }

        public int InvoiceNum { get; private set; }
        public DateTime Date { get; private set; }
        public string Customer { get; private set; }
        public string Company { get; private set; }
        public string Address { get; private set; }
        public string Address2 { get; private set; }
        public string Phone { get; private set; }
        public string Fax { get; private set; }
        public List<InvoiceData> Items { get { return items; } }
    }
}
