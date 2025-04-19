using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Interop;

namespace Win32_API_s
{
    internal class Program
    {

        //Change Desktop Wallpaper

        // Import the SystemParametersInfo function from user32.dll
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SystemParametersInfo(
            uint action, uint uParam, string vParam, uint winIni);

        // Constants for the function
        public static readonly uint SPI_SETDESKWALLPAPER = 0x14;
        public static readonly uint SPIF_UPDATEINIFILE = 0x01;
        public static readonly uint SPIF_SENDCHANGE = 0x02;


        public static void SetWallpaper(string path)
        {
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
            Console.WriteLine("Wallpaper changed successfully!");
        }

        /// <summary>
        /// Get Screen Resolution
        /// </summary>
        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);
        static void GetScreenResolution()
        {
            int screenWidth = GetSystemMetrics(0);  // SM_CXSCREEN = 0
            int screenHeight = GetSystemMetrics(1); // SM_CYSCREEN = 1

            Console.WriteLine("Screen Width: {0}, Screen Height: {1}", screenWidth, screenHeight);
        }
        /// <summary>
        /// Battery Info
        /// </summary>
        /// 

        // Define the SYSTEM_POWER_STATUS structure with fields as per the Windows API documentation
        [StructLayout(LayoutKind.Sequential)]
        public struct SYSTEM_POWER_STATUS
        {
            public byte ACLineStatus;
            public byte BatteryFlag;
            public byte BatteryLifePercent;
            public byte Reserved1;
            public int BatteryLifeTime;
            public int BatteryFullLifeTime;
        }

        // Import the GetSystemPowerStatus API from kernel32.dll
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool GetSystemPowerStatus(out SYSTEM_POWER_STATUS sps);

        static string GetBatteryStatus(byte flag)
        {
            switch (flag)
            {
                case 1:
                    return "High, more than 66% charged";
                case 2:
                    return "Low, less than 33% charged";
                case 4:
                    return "Critical, less than 5% charged";
                case 8:
                    return "Charging";
                case 128:
                    return "No battery";
                case 255:
                    return "Unknown status";
                default:
                    return "Battery status not detected";
            }
        }

        // Import MessageBox function from user32.dll
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern int MessageBox(IntPtr hWnd, String text, String caption, int type);

        //
        //Processes List
        [DllImport("psapi.dll", SetLastError = true)]
        public static extern bool EnumProcesses([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U4)][In][Out] uint[] processIds, uint arraySizeBytes, out uint bytesReturned);

        [DllImport("psapi.dll", SetLastError = true)]
        public static extern bool GetProcessMemoryInfo(IntPtr hProcess, out PROCESS_MEMORY_COUNTERS counters, uint size);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(uint processAccess, bool bInheritHandle, uint processId);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool CloseHandle(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_MEMORY_COUNTERS
        {
            public uint cb;
            public uint PageFaultCount;
            public uint PeakWorkingSetSize;
            public uint WorkingSetSize;
            public uint QuotaPeakPagedPoolUsage;
            public uint QuotaPagedPoolUsage;
            public uint QuotaPeakNonPagedPoolUsage;
            public uint QuotaNonPagedPoolUsage;
            public uint PagefileUsage;
            public uint PeakPagefileUsage;
        }

        const int PROCESS_QUERY_INFORMATION = 0x0400;
        const int PROCESS_VM_READ = 0x0010;



        public class WifiScanner
        {
            [DllImport("Wlanapi.dll")]
            private static extern uint WlanOpenHandle(uint dwClientVersion, IntPtr pReserved, out uint pdwNegotiatedVersion, out IntPtr phClientHandle);

            [DllImport("Wlanapi.dll")]
            private static extern uint WlanEnumInterfaces(IntPtr hClientHandle, IntPtr pReserved, out IntPtr ppInterfaceList);

            [DllImport("Wlanapi.dll")]
            private static extern uint WlanCloseHandle(IntPtr hClientHandle, IntPtr pReserved);


            [DllImport("Wlanapi.dll")]
            private static extern uint WlanFreeMemory(IntPtr pMemory);

            private IntPtr clientHandle = IntPtr.Zero;

            public WifiScanner()
            {
                uint negotiatedVersion;
                WlanOpenHandle(2, IntPtr.Zero, out negotiatedVersion, out clientHandle);
            }

            ~WifiScanner()
            {
                WlanCloseHandle(clientHandle, IntPtr.Zero);
            }

            public List<string> GetAvailableNetworks()
            {
                IntPtr interfaceList = IntPtr.Zero;
                WlanEnumInterfaces(clientHandle, IntPtr.Zero, out interfaceList);
                var listHeader = (WlanInterfaceInfoListHeader)Marshal.PtrToStructure(interfaceList, typeof(WlanInterfaceInfoListHeader));
                var wlanInterfaceInfo = new WlanInterfaceInfo[listHeader.dwNumberOfItems];
                List<string> networkList = new List<string>();

                for (int i = 0; i < listHeader.dwNumberOfItems; i++)
                {
                    IntPtr interfaceInfoPtr = new IntPtr(interfaceList.ToInt64() + (i * Marshal.SizeOf(typeof(WlanInterfaceInfo))) + Marshal.SizeOf(typeof(int)));
                    wlanInterfaceInfo[i] = (WlanInterfaceInfo)Marshal.PtrToStructure(interfaceInfoPtr, typeof(WlanInterfaceInfo));
                    networkList.Add(wlanInterfaceInfo[i].strProfileName);
                }

                WlanFreeMemory(interfaceList);
                return networkList;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct WlanInterfaceInfoListHeader
            {
                public uint dwNumberOfItems;
                public uint dwIndex;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            private struct WlanInterfaceInfo
            {
                public Guid InterfaceGuid;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
                public string strProfileName;
            }
        }


        //Send Email Via OutLook
        static void SendEmailViaOutLook()
        {
            try
            {
                Office.Outlook.Application outlookApp = new Office.Outlook.Application();
                Office.Outlook.MailItem mailItem = (Office.Outlook.MailItem)outlookApp.CreateItem(Office.Outlook.OlItemType.olMailItem);
                mailItem.Subject = "Test Email from C#";
                mailItem.To = "abdullah.h@alameensoft.com";  // Change this to the actual recipient's email address
                mailItem.Body = "Hello, this is a test email sent from a C# application using Outlook Interop.";
                mailItem.Importance = Office.Outlook.OlImportance.olImportanceHigh;
                mailItem.Display(false);  // Set to true to display the email before sending
                mailItem.Send();
                Console.WriteLine("Email sent successfully!");
                Console.ReadKey();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        
        static void ExcelSheetCreator() 
        {
            Office.Excel.Application excelApp = new Office.Excel.Application();
            try
            {
                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not properly installed!!");
                    return;
                }

                excelApp.Visible = true;  // Set to false to run Excel in the background

                // Create a new, empty workbook and add a worksheet
                Office.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                Office.Excel.Worksheet worksheet = (Office.Excel.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "MySheet";

                // Populate the worksheet with numbers 1 to 10
                for (int i = 1; i <= 10; i++)
                {
                    worksheet.Cells[i, 1] = i;
                    worksheet.Cells[i, 2] = "Item" + i.ToString();

                }

                // Save the workbook
                string filepath = @"C:\Temp\MyExcel.xlsx";  // Change the path as needed
                workbook.SaveAs(filepath);
                workbook.Close(true);
                Console.WriteLine("Excel file created successfully at: " + filepath);
                Console.ReadKey();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                excelApp.Quit();  // Close Excel application
            }

        }
        static void CreateawordDocument()
        {
            Office.Word.Application wordApp = new Office.Word.Application();
            try
            {
                wordApp.Visible = false;  // Set to true if you want to see Word while the document is being created

                Office.Word.Document doc = wordApp.Documents.Add();  // Create a new document
                Office.Word.Paragraph para = doc.Paragraphs.Add();   // Add a paragraph
                para.Range.Text = "Hi, My Name is Mohammed Abu-Hadhoud";  // Your name goes here

                // Save the document
                string filepath = @"C:\Temp\MyDocument.docx";  // Change the path as needed
                doc.SaveAs2(filepath);
                doc.Close();

                Console.WriteLine("Document created successfully at: " + filepath);
                Console.ReadKey();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                wordApp.Quit();  // Close Word application
            }
        }

        static void CreatePowerPoint() 
        {

            Office.PowerPoint.Application pptApplication = new Office.PowerPoint.Application();

            try
            {
                pptApplication.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                Office.PowerPoint.Presentations presentations = pptApplication.Presentations;
                Office.PowerPoint.Presentation presentation = presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

                // Add a slide
                Office.PowerPoint.Slides slides = presentation.Slides;
                Office.PowerPoint.Slide slide = slides.Add(1, Office.PowerPoint.PpSlideLayout.ppLayoutText);

                // Set title
                Office.PowerPoint.Shape titleShape = slide.Shapes[1];
                titleShape.TextFrame.TextRange.Text = "Hello, PowerPoint!";

                // Set subtitle
                Office.PowerPoint.Shape bodyShape = slide.Shapes[2];
                bodyShape.TextFrame.TextRange.Text = "Created using C#";

                // Save the presentation
                string filePath = @"C:\Temp\MyPresentation.pptx";
                presentation.SaveAs(filePath, Office.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);
                presentation.Close();
                Console.WriteLine("Presentation created successfully at: " + filePath);
                Console.ReadKey();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                pptApplication.Quit();
            }
        }


        static void Main()
        {

            CreateawordDocument();
            SendEmailViaOutLook();
            CreatePowerPoint();
            // The path to the wallpaper image
            string wallpaperPath = @"C:\pics\newpic.jpg";

            // Set the wallpaper
            //SetWallpaper(wallpaperPath);
            Console.ReadKey();

            if (GetSystemPowerStatus(out SYSTEM_POWER_STATUS status))
            {
                Console.WriteLine("Battery Information:");
                Console.WriteLine("AC Line Status: " + (status.ACLineStatus == 0 ? "Offline" : "Online"));
                Console.WriteLine("Battery Charge Status: " + GetBatteryStatus(status.BatteryFlag));
                Console.WriteLine("Battery Life Percent: " + (status.BatteryLifePercent == 255 ? "Unknown" : status.BatteryLifePercent + "%"));
                Console.WriteLine("Battery Life Remaining: " + (status.BatteryLifeTime == -1 ? "Unknown" : status.BatteryLifeTime + " seconds"));
                Console.WriteLine("Full Battery Lifetime: " + (status.BatteryFullLifeTime == -1 ? "Unknown" : status.BatteryFullLifeTime + " seconds"));
            }
            else
            {
                Console.WriteLine("Unable to get battery status.");
            }

            MessageBox(IntPtr.Zero, "Hello, World!", "My Message Box", 0);

            ///  PROCESS list
            uint[] processIds = new uint[1024];
            uint bytesReturned;

            if (EnumProcesses(processIds, (uint)processIds.Length * sizeof(uint), out bytesReturned))
            {
                Console.WriteLine("Number of processes: {0}", bytesReturned / sizeof(uint));

                for (int i = 0; i < bytesReturned / sizeof(uint); i++)
                {
                    uint pid = processIds[i];
                    IntPtr processHandle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, pid);

                    if (processHandle != IntPtr.Zero)
                    {
                        PROCESS_MEMORY_COUNTERS memCounters;
                        if (GetProcessMemoryInfo(processHandle, out memCounters, (uint)Marshal.SizeOf(typeof(PROCESS_MEMORY_COUNTERS))))
                        {
                            string processName = "Unknown";
                            try
                            {
                                Process proc = Process.GetProcessById((int)pid);
                                processName = proc.ProcessName;
                            }
                            catch (Exception)
                            {
                                // Process might have exited or access denied
                            }

                            Console.WriteLine($"Process ID: {pid}, Name: {processName} - Memory Usage: {memCounters.WorkingSetSize / 1024} KB");
                        }

                        CloseHandle(processHandle);
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to enumerate processes.");
            }


            var wifiScanner = new WifiScanner();
            var networks = wifiScanner.GetAvailableNetworks();
            foreach (var network in networks)
            {
                Console.WriteLine($"SSID: {network}");
            }
            Console.ReadKey();

        }
    }
}
