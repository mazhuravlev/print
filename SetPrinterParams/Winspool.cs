using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SetPrinterParams
{
    public static class Winspool
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct structPrinterDefaults
        {
            [MarshalAs(UnmanagedType.LPTStr)]
            public String pDatatype;
            public IntPtr pDevMode;
            [MarshalAs(UnmanagedType.I4)]
            public int DesiredAccess;
        };

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinter", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false, CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurity()]
        public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPTStr)]
            string printerName,
            out IntPtr phPrinter,
            ref structPrinterDefaults pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false,
             CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurity()]
        public static extern bool ClosePrinter(IntPtr phPrinter);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct structSize
        {
            public Int32 width;
            public Int32 height;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct structRect
        {
            public Int32 left;
            public Int32 top;
            public Int32 right;
            public Int32 bottom;
        }

        //[StructLayout(LayoutKind.Explicit, CharSet = CharSet.Unicode)]
        //public struct FormInfo1
        //{
        //    [FieldOffset(0), MarshalAs(UnmanagedType.I4)]
        //    public uint Flags;
        //    [FieldOffset(4), MarshalAs(UnmanagedType.LPWStr)]
        //    public String pName;
        //    [FieldOffset(8)]
        //    public structSize Size;
        //    [FieldOffset(16)]
        //    public structRect ImageableArea;
        //};
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Size = 32)]
        public struct FormInfo1
        {
            [MarshalAs(UnmanagedType.I4)]
            public uint Flags;
            [MarshalAs(UnmanagedType.LPWStr)]
            public String pName;
            public structSize Size;
            public structRect ImageableArea;
        };
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi/* changed from CharSet=CharSet.Auto */)]
        public struct structDevMode
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public String
                dmDeviceName;
            [MarshalAs(UnmanagedType.U2)]
            public short dmSpecVersion;
            [MarshalAs(UnmanagedType.U2)]
            public short dmDriverVersion;
            [MarshalAs(UnmanagedType.U2)]
            public short dmSize;
            [MarshalAs(UnmanagedType.U2)]
            public short dmDriverExtra;
            [MarshalAs(UnmanagedType.U4)]
            public int dmFields;
            [MarshalAs(UnmanagedType.I2)]
            public short dmOrientation;
            [MarshalAs(UnmanagedType.I2)]
            public short dmPaperSize;
            [MarshalAs(UnmanagedType.I2)]
            public short dmPaperLength;
            [MarshalAs(UnmanagedType.I2)]
            public short dmPaperWidth;
            [MarshalAs(UnmanagedType.I2)]
            public short dmScale;
            [MarshalAs(UnmanagedType.I2)]
            public short dmCopies;
            [MarshalAs(UnmanagedType.I2)]
            public short dmDefaultSource;
            [MarshalAs(UnmanagedType.I2)]
            public short dmPrintQuality;
            [MarshalAs(UnmanagedType.I2)]
            public short dmColor;
            [MarshalAs(UnmanagedType.I2)]
            public short dmDuplex;
            [MarshalAs(UnmanagedType.I2)]
            public short dmYResolution;
            [MarshalAs(UnmanagedType.I2)]
            public short dmTTOption;
            [MarshalAs(UnmanagedType.I2)]
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public String dmFormName;
            [MarshalAs(UnmanagedType.U2)]
            public short dmLogPixels;
            [MarshalAs(UnmanagedType.U4)]
            public int dmBitsPerPel;
            [MarshalAs(UnmanagedType.U4)]
            public int dmPelsWidth;
            [MarshalAs(UnmanagedType.U4)]
            public int dmPelsHeight;
            [MarshalAs(UnmanagedType.U4)]
            public int dmNup;
            [MarshalAs(UnmanagedType.U4)]
            public int dmDisplayFrequency;
            [MarshalAs(UnmanagedType.U4)]
            public int dmICMMethod;
            [MarshalAs(UnmanagedType.U4)]
            public int dmICMIntent;
            [MarshalAs(UnmanagedType.U4)]
            public int dmMediaType;
            [MarshalAs(UnmanagedType.U4)]
            public int dmDitherType;
            [MarshalAs(UnmanagedType.U4)]
            public int dmReserved1;
            [MarshalAs(UnmanagedType.U4)]
            public int dmReserved2;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct PRINTER_INFO_9
        {
            public IntPtr pDevMode;
        }

        [DllImport("winspool.Drv", EntryPoint = "AddFormW", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = true,
             CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurity()]
        public static extern bool AddForm(
         IntPtr phPrinter,
            [MarshalAs(UnmanagedType.I4)] int level,
         ref FormInfo1 form);

        /*    This method is not used
                [DllImport("winspool.Drv", EntryPoint="SetForm", SetLastError=true,
                     CharSet=CharSet.Unicode, ExactSpelling=false,
                     CallingConvention=CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
                public static extern bool SetForm(IntPtr phPrinter, string paperName,
                    [MarshalAs(UnmanagedType.I4)] int level, ref FormInfo1 form);
        */
        [DllImport("winspool.Drv", EntryPoint = "DeleteForm", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false, CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurity()]
        public static extern bool DeleteForm(
         IntPtr phPrinter,
            [MarshalAs(UnmanagedType.LPTStr)] string pName);

        [DllImport("kernel32.dll", EntryPoint = "GetLastError", SetLastError = false,
             ExactSpelling = true, CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurity()]
        public static extern Int32 GetLastError();

        [DllImport("GDI32.dll", EntryPoint = "CreateDC", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false,
             CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurity()]
        public static extern IntPtr CreateDC([MarshalAs(UnmanagedType.LPTStr)]
            string pDrive,
            [MarshalAs(UnmanagedType.LPTStr)] string pName,
            [MarshalAs(UnmanagedType.LPTStr)] string pOutput,
            ref structDevMode pDevMode);

        [DllImport("GDI32.dll", EntryPoint = "ResetDC", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false,
             CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurity()]
        public static extern IntPtr ResetDC(
         IntPtr hDC,
         ref structDevMode
            pDevMode);

        [DllImport("GDI32.dll", EntryPoint = "DeleteDC", SetLastError = true,
             CharSet = CharSet.Unicode, ExactSpelling = false,
             CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurity()]
        public static extern bool DeleteDC(IntPtr hDC);

        [DllImport("winspool.Drv", EntryPoint = "SetPrinterA", SetLastError = true,
            CharSet = CharSet.Auto, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurity()]
        public static extern bool SetPrinter(
           IntPtr hPrinter,
           [MarshalAs(UnmanagedType.I4)] int level,
           IntPtr pPrinter,
           [MarshalAs(UnmanagedType.I4)] int command);

        /*
         LONG DocumentProperties(
           HWND hWnd,               // handle to parent window 
           HANDLE hPrinter,         // handle to printer object
           LPTSTR pDeviceName,      // device name
           PDEVMODE pDevModeOutput, // modified device mode
           PDEVMODE pDevModeInput,  // original device mode
           DWORD fMode              // mode options
           );
         */
        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesA", SetLastError = true,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern int DocumentProperties(
           IntPtr hwnd,
           IntPtr hPrinter,
           [MarshalAs(UnmanagedType.LPStr)] string pDeviceName /* changed from String to string */,
           IntPtr pDevModeOutput,
           IntPtr pDevModeInput,
           int fMode
           );

        [DllImport("winspool.Drv", EntryPoint = "GetPrinterA", SetLastError = true,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool GetPrinter(
           IntPtr hPrinter,
           int dwLevel /* changed type from Int32 */,
           IntPtr pPrinter,
           int dwBuf /* chagned from Int32*/,
           out int dwNeeded /* changed from Int32*/
           );

        // SendMessageTimeout tools
        [Flags]
        public enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0000,
            SMTO_BLOCK = 0x0001,
            SMTO_ABORTIFHUNG = 0x0002,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x0008
        }
        public const int WM_SETTINGCHANGE = 0x001A;
        public const int HWND_BROADCAST = 0xffff;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
           IntPtr windowHandle,
           uint Msg,
           IntPtr wParam,
           IntPtr lParam,
           SendMessageTimeoutFlags flags,
           uint timeout,
           out IntPtr result
           );

        [DllImport("winspool.Drv", EntryPoint = "SetForm", SetLastError = true,
        CharSet = CharSet.Unicode, ExactSpelling = false,
        CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        public static extern bool SetForm(IntPtr phPrinter, string paperName,
       [MarshalAs(UnmanagedType.I4)] int level, ref FormInfo1 form);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool EnumPrinters(PrinterEnumFlags Flags, string Name, uint Level, IntPtr pPrinterEnum, uint cbBuf, ref uint pcbNeeded, ref uint pcReturned);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool EnumForms(IntPtr phPrinter, uint Level, IntPtr pFormsEnum, uint cbBuf, ref uint pcbNeeded, ref uint pcReturned);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct PRINTER_INFO_2
        {
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pServerName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pPrinterName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pShareName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pPortName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pDriverName;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pComment;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pLocation;
            public IntPtr pDevMode;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pSepFile;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pPrintProcessor;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pDatatype;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pParameters;
            public IntPtr pSecurityDescriptor;
            public uint Attributes; // See note below!
            public uint Priority;
            public uint DefaultPriority;
            public uint StartTime;
            public uint UntilTime;
            public uint Status;
            public uint cJobs;
            public uint AveragePPM;
        }

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);
        
        [FlagsAttribute]
        public enum PrinterEnumFlags
        {
            PRINTER_ENUM_DEFAULT = 0x00000001,
            PRINTER_ENUM_LOCAL = 0x00000002,
            PRINTER_ENUM_CONNECTIONS = 0x00000004,
            PRINTER_ENUM_FAVORITE = 0x00000004,
            PRINTER_ENUM_NAME = 0x00000008,
            PRINTER_ENUM_REMOTE = 0x00000010,
            PRINTER_ENUM_SHARED = 0x00000020,
            PRINTER_ENUM_NETWORK = 0x00000040,
            PRINTER_ENUM_EXPAND = 0x00004000,
            PRINTER_ENUM_CONTAINER = 0x00008000,
            PRINTER_ENUM_ICONMASK = 0x00ff0000,
            PRINTER_ENUM_ICON1 = 0x00010000,
            PRINTER_ENUM_ICON2 = 0x00020000,
            PRINTER_ENUM_ICON3 = 0x00040000,
            PRINTER_ENUM_ICON4 = 0x00080000,
            PRINTER_ENUM_ICON5 = 0x00100000,
            PRINTER_ENUM_ICON6 = 0x00200000,
            PRINTER_ENUM_ICON7 = 0x00400000,
            PRINTER_ENUM_ICON8 = 0x00800000,
            PRINTER_ENUM_HIDE = 0x01000000,
            PRINTER_ENUM_CATEGORY_ALL = 0x02000000,
            PRINTER_ENUM_CATEGORY_3D = 0x04000000
        }
    }
}
