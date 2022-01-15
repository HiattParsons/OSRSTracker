using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Timers;
using System.Net;
using System.Diagnostics;

//downloads OSRS highscores and saves them to an Excel spreadsheet
namespace OSRSTracker
{
    class Program
    {
        private static System.Timers.Timer aTimer;
        static void Main(string[] args)
        {
            //starts timer and gives starting information
            Console.SetWindowSize(40, 20);
            Console.WriteLine("OSRS highscore tracker : zesty125 Ironman scores\n");
            Console.WriteLine("currently set to : 1 hour interval");
            SetTimer();
            Console.WriteLine("press enter to exit (haha)");
            return;
            Console.ReadLine();

        }

        //sets a timer for 1 hour and resets timer when it completes
        private static void SetTimer()
        {
            aTimer = new System.Timers.Timer(3600000);
            aTimer.Elapsed += OnTimedEvent;
            aTimer.AutoReset = true;
            aTimer.Enabled = true;
        }

        //triggered when timer completes
        private static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            //creates a webclient that retrives a string of JSON, then splits the string up 
            string HSData;
            char[] delimiterChars = { ' ', ',', '\n' };
            Console.WriteLine("\naccessing high scores");
            WebClient client = new WebClient();
            HSData = client.DownloadString("https://secure.runescape.com/m=hiscore_oldschool_ironman/index_lite.ws?player=zesty125");
            Console.WriteLine("scores are obtained");
            string[] HSSplit = HSData.Split(delimiterChars);

            //Opens Excel
            Console.WriteLine("opening excel");
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;           
            xlApp = new Excel.Application();
            
            //Selects the Excel file 
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\hiatt\Desktop\OSRS tracker\tracker2\OSRSHS_zesty125.xlsx", 0,
                false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //selects the range of cells to be written to 
            range = xlWorkSheet.Cells[1, 1];
            double rwS = range.Value;
            int rw = (int)(rwS + 1);
            int cl = 192;

            //iterates through the cells and places the appropriate value 

            Console.WriteLine("writing to excel");
            for (int i = 2; i < cl; i++)
            {
                xlWorkSheet.Cells[rw, i] = HSSplit[i - 2];
            }
            
            //places timestamp and increases the iteration counter
            DateTime now = DateTime.Now;
            xlWorkSheet.Cells[rw, 1] = now;
            xlWorkSheet.Cells[1, 1] = rw;

            //saves and quits Excel
            xlWorkBook.Save();
            xlWorkBook.Close();
            xlApp.Quit();
            Console.WriteLine("time: " + now);
            Console.WriteLine("excel is saved and closed");
            Console.WriteLine("timer is reset\n");
        }
    }
}
