using System;
using System.Collections.Generic;
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

        }
    }
}
