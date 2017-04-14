using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Pdf;

namespace Pd
{
    class Program
    {
        static void Main(string[] args)
        {
            PdfDocument doc = new PdfDocument();
            doc.LoadFromFile("sample.pdf");

            //Use the default printer to print all the pages
            //doc.PrintDocument.Print();

            //Set the printer and select the pages you want to print
            PrintDocument printDoc = doc.PrintDocument;

            printDoc.DefaultPageSettings.PaperSize = new System.Drawing.Printing.PaperSize("Test", 200, 200);
            printDoc.PrinterSettings.PrinterName = "\\\\prn02pik.picompany.ru\\10-163";
            //printDoc.PrinterSettings.PrinterName = "Adobe PDF";
            doc.PrintFromPage = 1;
            doc.PrintToPage = 1;

            printDoc.Print();

            /* PrintDialog dialogPrint = new PrintDialog();
             dialogPrint.AllowPrintToFile = true;
             dialogPrint.AllowSomePages = true;
             dialogPrint.PrinterSettings.MinimumPage = 1;
             dialogPrint.PrinterSettings.MaximumPage = doc.Pages.Count;
             dialogPrint.PrinterSettings.FromPage = 1;
             dialogPrint.PrinterSettings.ToPage = doc.Pages.Count;

             if (dialogPrint.ShowDialog() == DialogResult.OK)
             {
                 doc.PrintFromPage = dialogPrint.PrinterSettings.FromPage;
                 doc.PrintToPage = dialogPrint.PrinterSettings.ToPage;
                 doc.PrinterName = dialogPrint.PrinterSettings.PrinterName;


                 dialogPrint.Document = printDoc;
                 printDoc.Print();
             }*/
        }
    }
}
