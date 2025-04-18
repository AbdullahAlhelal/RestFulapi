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


        static void Main()
        {
            // The path to the wallpaper image
            string wallpaperPath = @"C:\pics\newpic.jpg";

            // Set the wallpaper
            SetWallpaper(wallpaperPath);
            Console.ReadKey();

        }
    }
}
