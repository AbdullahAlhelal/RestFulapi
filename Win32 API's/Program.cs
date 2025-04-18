using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Win32_API_s
{
    internal class Program
    {

        //Change Desktop Wallpaper

        // Import the SystemParametersInfo function from user32.dll
        [DllImport("user32.dll" , CharSet = CharSet.Auto)]
        public static extern int SystemParametersInfo(
            uint action , uint uParam , string vParam , uint winIni);

        // Constants for the function
        public static readonly uint SPI_SETDESKWALLPAPER = 0x14;
        public static readonly uint SPIF_UPDATEINIFILE = 0x01;
        public static readonly uint SPIF_SENDCHANGE = 0x02;

 
        public static void SetWallpaper(string path)
        {
            SystemParametersInfo(SPI_SETDESKWALLPAPER , 0 , path , SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
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

            Console.WriteLine("Screen Width: {0}, Screen Height: {1}" , screenWidth , screenHeight);
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
        [DllImport("kernel32.dll" , SetLastError = true)]
        static extern bool GetSystemPowerStatus(out SYSTEM_POWER_STATUS sps);
          
            static string GetBatteryStatus(byte flag)
        {
            switch ( flag )
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
        [DllImport("user32.dll" , CharSet = CharSet.Unicode , SetLastError = true)]
        static extern int MessageBox(IntPtr hWnd , String text , String caption , int type);

        //
        //Processes List
       [DllImport("psapi.dll" , SetLastError = true)]
        public static extern bool EnumProcesses([MarshalAs(UnmanagedType.LPArray , ArraySubType = UnmanagedType.U4)][In][Out] uint[] processIds , uint arraySizeBytes , out uint bytesReturned);

        [DllImport("psapi.dll" , SetLastError = true)]
        public static extern bool GetProcessMemoryInfo(IntPtr hProcess , out PROCESS_MEMORY_COUNTERS counters , uint size);

        [DllImport("kernel32.dll" , SetLastError = true)]
        public static extern IntPtr OpenProcess(uint processAccess , bool bInheritHandle , uint processId);

        [DllImport("kernel32.dll" , SetLastError = true)]
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
            private static extern uint WlanOpenHandle(uint dwClientVersion , IntPtr pReserved , out uint pdwNegotiatedVersion , out IntPtr phClientHandle);

            [DllImport("Wlanapi.dll")]
            private static extern uint WlanEnumInterfaces(IntPtr hClientHandle , IntPtr pReserved , out IntPtr ppInterfaceList);

            [DllImport("Wlanapi.dll")]
            private static extern uint WlanCloseHandle(IntPtr hClientHandle , IntPtr pReserved);

            [DllImport("Wlanapi.dll")]
            private static extern uint WlanFreeMemory(IntPtr pMemory);

            private IntPtr clientHandle = IntPtr.Zero;

            public WifiScanner()
            {
                uint negotiatedVersion;
                WlanOpenHandle(2 , IntPtr.Zero , out negotiatedVersion , out clientHandle);
            }

            ~WifiScanner()
            {
                WlanCloseHandle(clientHandle , IntPtr.Zero);
            }

            public List<string> GetAvailableNetworks()
            {
                IntPtr interfaceList = IntPtr.Zero;
                WlanEnumInterfaces(clientHandle , IntPtr.Zero , out interfaceList);
                var listHeader = (WlanInterfaceInfoListHeader) Marshal.PtrToStructure(interfaceList , typeof(WlanInterfaceInfoListHeader));
                var wlanInterfaceInfo = new WlanInterfaceInfo[listHeader.dwNumberOfItems];
                List<string> networkList = new List<string>();

                for ( int i = 0 ; i < listHeader.dwNumberOfItems ; i++ )
                {
                    IntPtr interfaceInfoPtr = new IntPtr(interfaceList.ToInt64() + (i * Marshal.SizeOf(typeof(WlanInterfaceInfo))) + Marshal.SizeOf(typeof(int)));
                    wlanInterfaceInfo[i] = (WlanInterfaceInfo) Marshal.PtrToStructure(interfaceInfoPtr , typeof(WlanInterfaceInfo));
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

            [StructLayout(LayoutKind.Sequential , CharSet = CharSet.Unicode)]
            private struct WlanInterfaceInfo
            {
                public Guid InterfaceGuid;
                [MarshalAs(UnmanagedType.ByValTStr , SizeConst = 256)]
                public string strProfileName;
            }
        }

        static void Main()
        {
            // The path to the wallpaper image
            string wallpaperPath = @"C:\pics\newpic.jpg";

            // Set the wallpaper
            //SetWallpaper(wallpaperPath);
            Console.ReadKey();

                if ( GetSystemPowerStatus(out SYSTEM_POWER_STATUS status) )
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

            MessageBox(IntPtr.Zero , "Hello, World!" , "My Message Box" , 0);

            ///  PROCESS list
            uint[] processIds = new uint[1024];
            uint bytesReturned;

            if ( EnumProcesses(processIds , (uint) processIds.Length * sizeof(uint) , out bytesReturned) )
            {
                Console.WriteLine("Number of processes: {0}" , bytesReturned / sizeof(uint));

                for ( int i = 0 ; i < bytesReturned / sizeof(uint) ; i++ )
                {
                    uint pid = processIds[i];
                    IntPtr processHandle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ , false , pid);

                    if ( processHandle != IntPtr.Zero )
                    {
                        PROCESS_MEMORY_COUNTERS memCounters;
                        if ( GetProcessMemoryInfo(processHandle , out memCounters , (uint) Marshal.SizeOf(typeof(PROCESS_MEMORY_COUNTERS))) )
                        {
                            string processName = "Unknown";
                            try
                            {
                                Process proc = Process.GetProcessById((int) pid);
                                processName = proc.ProcessName;
                            }
                            catch ( Exception )
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
            foreach ( var network in networks )
            {
                Console.WriteLine($"SSID: {network}");
            }
            Console.ReadKey();

        }
    }
}
