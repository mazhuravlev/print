using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Printing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SetPrinterParams
{
    public static class PrinterUtils
    {
        /// <summary>
        /// Add the printer form to a printer 
        /// </summary>
        /// <param name="printerName">The printer name</param>
        /// <param name="paperName">Name of the printer form</param>
        /// <param name="widthMm">Width given in millimeters</param>
        /// <param name="heightMm">Height given in millimeters</param>
        public static void AddCustomPaperSize(string printerName, string paperName, float
            widthMm, float heightMm)
        {
            if (PlatformID.Win32NT == Environment.OSVersion.Platform)
            {
                // The code to add a custom paper size is different for Windows NT then it is
                // for previous versions of windows

                const int PRINTER_ACCESS_USE = 0x00000008;
                const int PRINTER_ACCESS_ADMINISTER = 0x00000004;
                const int FORM_PRINTER = 0x00000002;

                var defaults = new Winspool.structPrinterDefaults();
                defaults.pDatatype = null;
                defaults.pDevMode = IntPtr.Zero;
                defaults.DesiredAccess = PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE;

                var hPrinter = IntPtr.Zero;

                // Open the printer.
                if (Winspool.OpenPrinter(printerName, out hPrinter, ref defaults))
                {
                    try
                    {
                        // delete the form incase it already exists
                        Winspool.DeleteForm(hPrinter, paperName);
                        // create and initialize the FORM_INFO_1 structure
                        var formInfo = new Winspool.FormInfo1();
                        formInfo.Flags = 0;
                        formInfo.pName = paperName;
                        // all sizes in 1000ths of millimeters
                        formInfo.Size.width = (int)(widthMm * 1000.0);
                        formInfo.Size.height = (int)(heightMm * 1000.0);
                        formInfo.ImageableArea.left = 0;
                        formInfo.ImageableArea.right = formInfo.Size.width;
                        formInfo.ImageableArea.top = 0;
                        formInfo.ImageableArea.bottom = formInfo.Size.height;
                        if (!Winspool.AddForm(hPrinter, 1, ref formInfo))
                        {
                            var strBuilder = new StringBuilder();
                            strBuilder.AppendFormat("Failed to add the custom paper size {0} to the printer {1}, System error number: {2}",
                                paperName, printerName, Winspool.GetLastError());
                            throw new ApplicationException(strBuilder.ToString());
                        }

                        // INIT
                        const int DM_OUT_BUFFER = 2;
                        const int DM_IN_BUFFER = 8;
                        var devMode = new Winspool.structDevMode();
                        IntPtr hPrinterInfo, hDummy;
                        Winspool.PRINTER_INFO_9 printerInfo;
                        printerInfo.pDevMode = IntPtr.Zero;
                        int iPrinterInfoSize, iDummyInt;


                        // GET THE SIZE OF THE DEV_MODE BUFFER
                        var iDevModeSize = Winspool.DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);

                        if (iDevModeSize < 0)
                            throw new ApplicationException("Cannot get the size of the DEVMODE structure.");

                        // ALLOCATE THE BUFFER
                        var hDevMode = Marshal.AllocCoTaskMem(iDevModeSize + 100);

                        // GET A POINTER TO THE DEV_MODE BUFFER 
                        var iRet = Winspool.DocumentProperties(IntPtr.Zero, hPrinter, printerName, hDevMode, IntPtr.Zero, DM_OUT_BUFFER);

                        if (iRet < 0)
                            throw new ApplicationException("Cannot get the DEVMODE structure.");

                        // FILL THE DEV_MODE STRUCTURE
                        devMode = (Winspool.structDevMode)Marshal.PtrToStructure(hDevMode, devMode.GetType());

                        // SET THE FORM NAME FIELDS TO INDICATE THAT THIS FIELD WILL BE MODIFIED
                        devMode.dmFields = 0x10000; // DM_FORMNAME 
                        // SET THE FORM NAME
                        devMode.dmFormName = paperName;

                        // PUT THE DEV_MODE STRUCTURE BACK INTO THE POINTER
                        Marshal.StructureToPtr(devMode, hDevMode, true);

                        // MERGE THE NEW CHAGES WITH THE OLD
                        iRet = Winspool.DocumentProperties(IntPtr.Zero, hPrinter, printerName,
                                 printerInfo.pDevMode, printerInfo.pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);

                        if (iRet < 0)
                            throw new ApplicationException("Unable to set the orientation setting for this printer.");

                        // GET THE PRINTER INFO SIZE
                        Winspool.GetPrinter(hPrinter, 9, IntPtr.Zero, 0, out iPrinterInfoSize);
                        if (iPrinterInfoSize == 0)
                            throw new ApplicationException("GetPrinter failed. Couldn't get the # bytes needed for shared PRINTER_INFO_9 structure");

                        // ALLOCATE THE BUFFER
                        hPrinterInfo = Marshal.AllocCoTaskMem(iPrinterInfoSize + 100);

                        // GET A POINTER TO THE PRINTER INFO BUFFER
                        var bSuccess = Winspool.GetPrinter(hPrinter, 9, hPrinterInfo, iPrinterInfoSize, out iDummyInt);

                        if (!bSuccess)
                            throw new ApplicationException("GetPrinter failed. Couldn't get the shared PRINTER_INFO_9 structure");

                        // FILL THE PRINTER INFO STRUCTURE
                        printerInfo = (Winspool.PRINTER_INFO_9)Marshal.PtrToStructure(hPrinterInfo, printerInfo.GetType());
                        printerInfo.pDevMode = hDevMode;

                        // GET A POINTER TO THE PRINTER INFO STRUCTURE
                        Marshal.StructureToPtr(printerInfo, hPrinterInfo, true);

                        // SET THE PRINTER SETTINGS
                        bSuccess = Winspool.SetPrinter(hPrinter, 9, hPrinterInfo, 0);

                        if (!bSuccess)
                            throw new Win32Exception(Marshal.GetLastWin32Error(), "SetPrinter() failed.  Couldn't set the printer settings");

                        // Tell all open programs that this change occurred.
                        Winspool.SendMessageTimeout(
                           new IntPtr(Winspool.HWND_BROADCAST),
                           Winspool.WM_SETTINGCHANGE,
                           IntPtr.Zero,
                           IntPtr.Zero,
                           Winspool.SendMessageTimeoutFlags.SMTO_NORMAL,
                           1000,
                           out hDummy);
                    }
                    finally
                    {
                        Winspool.ClosePrinter(hPrinter);
                    }
                }
                else
                {
                    var strBuilder = new StringBuilder();
                    strBuilder.AppendFormat("Failed to open the {0} printer, System error number: {1}",
                        printerName, Winspool.GetLastError());
                    throw new ApplicationException(strBuilder.ToString());
                }
            }
            else
            {
                var pDevMode = new Winspool.structDevMode();
                var hDC = Winspool.CreateDC(null, printerName, null, ref pDevMode);
                if (hDC != IntPtr.Zero)
                {
                    const long DM_PAPERSIZE = 0x00000002L;
                    const long DM_PAPERLENGTH = 0x00000004L;
                    const long DM_PAPERWIDTH = 0x00000008L;
                    pDevMode.dmFields = (int)(DM_PAPERSIZE | DM_PAPERWIDTH | DM_PAPERLENGTH);
                    pDevMode.dmPaperSize = 256;
                    pDevMode.dmPaperWidth = (short)(widthMm * 1000.0);
                    pDevMode.dmPaperLength = (short)(heightMm * 1000.0);
                    Winspool.ResetDC(hDC, ref pDevMode);
                    Winspool.DeleteDC(hDC);
                }
            }
        }

        /// <summary>
        /// Adds the printer form to the default printer
        /// </summary>
        /// <param name="paperName">Name of the printer form</param>
        /// <param name="widthMm">Width given in millimeters</param>
        /// <param name="heightMm">Height given in millimeters</param>
        public static int AddCustomPaperSizeToDefaultPrinter(string paperName, float widthMm, float heightMm)
        {
            var pd = new PrintDocument();
            var sPrinterName = pd.PrinterSettings.PrinterName;
            AddCustomPaperSize(sPrinterName, paperName, widthMm, heightMm);
            return 0;
        }

        private const int ERROR_INSUFFICIENT_BUFFER = 122;

        public static Winspool.PRINTER_INFO_2[] EnumPrinters(Winspool.PrinterEnumFlags Flags)
        {
            uint cbNeeded = 0;
            uint cReturned = 0;
            if (Winspool.EnumPrinters(Flags, null, 2, IntPtr.Zero, 0, ref cbNeeded, ref cReturned))
            {
                return null;
            }
            int lastWin32Error = Marshal.GetLastWin32Error();
            if (lastWin32Error == ERROR_INSUFFICIENT_BUFFER)
            {
                IntPtr pAddr = Marshal.AllocHGlobal((int)cbNeeded);
                if (Winspool.EnumPrinters(Flags, null, 2, pAddr, cbNeeded, ref cbNeeded, ref cReturned))
                {
                    Winspool.PRINTER_INFO_2[] printerInfo2 = new Winspool.PRINTER_INFO_2[cReturned];
                    int offset = pAddr.ToInt32();
                    Type type = typeof(Winspool.PRINTER_INFO_2);
                    int increment = Marshal.SizeOf(type);
                    for (int i = 0; i < cReturned; i++)
                    {
                        printerInfo2[i] = (Winspool.PRINTER_INFO_2)Marshal.PtrToStructure(new IntPtr(offset), type);
                        offset += increment;
                    }
                    Marshal.FreeHGlobal(pAddr);
                    return printerInfo2;
                }
                lastWin32Error = Marshal.GetLastWin32Error();
            }
            throw new Win32Exception(lastWin32Error);
        }

        public static Winspool.FormInfo1[] EnumForms(IntPtr hPrinter)
        {
            uint cbNeeded = 0;
            uint cReturned = 0;
            if (Winspool.EnumForms(hPrinter, 1, IntPtr.Zero, 0, ref cbNeeded, ref cReturned))
            {
                return null;
            }
            int lastWin32Error = Marshal.GetLastWin32Error();
            if (lastWin32Error == ERROR_INSUFFICIENT_BUFFER)
            {
                IntPtr pAddr = Marshal.AllocHGlobal((int)cbNeeded);
                if (Winspool.EnumForms(hPrinter, 1, pAddr, cbNeeded, ref cbNeeded, ref cReturned))
                {
                    Winspool.FormInfo1[] formInfo = new Winspool.FormInfo1[cReturned];
                    int offset = pAddr.ToInt32();
                    Type type = typeof(Winspool.FormInfo1);
                    int increment = Marshal.SizeOf(type);
                    for (int i = 0; i < cReturned; i++)
                    {
                        formInfo[i] = (Winspool.FormInfo1)Marshal.PtrToStructure(new IntPtr(offset), type);
                        offset += increment;
                    }
                    Marshal.FreeHGlobal(pAddr);
                    return formInfo;
                }
                lastWin32Error = Marshal.GetLastWin32Error();
            }
            throw new Win32Exception(lastWin32Error);
        }
    }
}
