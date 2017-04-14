using System;
using System.Linq;

namespace SetPrinterParams
{
    public class PrinterUtilsWin
    {
        public static void Main()
        {
            const string printerName = "\\\\prn02pik.picompany.ru\\10-163";
            //const string printerName = "Adobe PDF";
            const string paperName = "AUTOFORMAT";
            IntPtr hPrinter;
            const int PRINTER_ACCESS_USE = 0x00000008;
            const int PRINTER_ACCESS_ADMINISTER = 0x00000004;
            const int STANDARD_RIGHTS_REQUIRED = 0x000F0000;
            const int PRINTER_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED | PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE;
            const int FORM_PRINTER = 0x00000002;

            var defaults = new Winspool.structPrinterDefaults
            {
                pDatatype = null,
                pDevMode = IntPtr.Zero,
                DesiredAccess = PRINTER_ALL_ACCESS //PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE
            };

            var printers = PrinterUtils.EnumPrinters(Winspool.PrinterEnumFlags.PRINTER_ENUM_LOCAL |
                                          Winspool.PrinterEnumFlags.PRINTER_ENUM_NETWORK |
                                          Winspool.PrinterEnumFlags.PRINTER_ENUM_DEFAULT | 
                                          Winspool.PrinterEnumFlags.PRINTER_ENUM_CONNECTIONS |
                                          Winspool.PrinterEnumFlags.PRINTER_ENUM_HIDE |
                                          Winspool.PrinterEnumFlags.PRINTER_ENUM_REMOTE);

            if(!Winspool.OpenPrinter(printerName, out hPrinter, ref defaults))
            {
                var error = Winspool.GetLastError();
                throw new Exception("Unable to open printer");
            }

            //Winspool.DeleteForm(hPrinter, paperName);

            var formInfo = new Winspool.FormInfo1
            {
                Flags = 0,
                pName = paperName,
                Size =
                {
                    width = 10000,
                    height = 10000
                },
                ImageableArea = { left = 0 }
            };
            formInfo.ImageableArea.right = formInfo.Size.width;
            formInfo.ImageableArea.top = 0;
            formInfo.ImageableArea.bottom = formInfo.Size.height;

            /*var deletedCount = 0;
            foreach (var formName in PrinterUtils.EnumForms(hPrinter).Select(x => x.pName))
            {
                if (Winspool.DeleteForm(hPrinter, formName)) deletedCount++;
            }*/

            Winspool.DeleteForm(hPrinter, paperName);

            
            if (!Winspool.AddForm(hPrinter, 1, ref formInfo))
            {
                var error = Winspool.GetLastError();
                Winspool.ClosePrinter(hPrinter);
                Console.WriteLine("Failed to add form");
            }
            

            if (!Winspool.SetForm(hPrinter, paperName, 1, ref formInfo))
            {
                var error = Winspool.GetLastError();
                Console.WriteLine("Failed to set form");
            }

            try
            {
                var forms = PrinterUtils.EnumForms(hPrinter).Select(x => x.pName).ToList();

                foreach (var form in forms)
                {
                    Console.WriteLine(form);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to enumerate forms: " + e.Message);
            }

            Winspool.ClosePrinter(hPrinter);

            if (!Winspool.SetDefaultPrinter(printerName))
            {
                throw new Exception("Failed to set default printer");
            }
            Console.WriteLine("OK");
            Console.ReadLine();
        }
    }
}
