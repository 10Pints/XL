using Spire.Xls;

using XL_Lib;
using static XL_Lib.Excel;

/*
using Spire;
using Spire.Xls;
using Spire.Xls.Collections;

using System.Collections.Generic;
using System.Text;
using System.Web;
using System.Xml.Xsl;
*/

namespace XL
{
   /// <summary>
   /// https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/How-to-Convert-Excel-to-Text-in-C-and-VB.NET.html
   /// </summary>
   public class Program
   {
      //static IConfigurationRoot? Config;
      /// <summary>
      /// Saves the specified workbook/worksheet as a tab separated text file
      /// </summary>
      /// <param name="args[0]">Optional settings file, default: XL_settings.json</param>
      /// 
      /// Note:
      /// if the config for a worbook has worksheets= '* 'then saves all worksheets in that workbook
      /// <exception cref="Exception"></exception>
      ///
      /// Args: 
      /// Preconditions:
      /// None ?
      /// 
      /// Postconditions:
      /// Returns 0 if success
      /// else  -1 if uninitialised (code error)
      /// else  100 and error msg displayed on the consoles
      /// if exception raised 

      public static int Main(string[] args)
      {
         int ret = -1; // Error 

         try
         { 
            Console.WriteLine($"XL 000: starting");
            Dictionary<string, Dictionary<string, string?>> map = Init(args);
            Workbook tmpWorkbook = new Workbook();

            foreach ( var cfg in map)
            {
               if (cfg.Value["worksheetTabNm"] == "*")
                  SaveWorkbook(cfg.Value);
               else
                  SaveWorksheet( cfg.Value, tmpWorkbook);
            }

            // Finally
            ret = 0; // Success
         }
         catch(Exception e)
         {
            Console.WriteLine($"XL Error 520: caught exception {e}");
            ret = 100;
         }

         Console.WriteLine($"XL 999: leaving, ret: {ret}");
         return ret;
      }
   }
}
