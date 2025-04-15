using Microsoft.Extensions.Configuration;
using Microsoft.VisualBasic;

using Spire.Xls;
using System;
using Spire.Xls.Core;

using System.Text;

namespace XL_Lib
{
   public static class Excel
   {
      /// <summary>
      /// RETURNS a list of configurations for the SaveWorksheet routine.
      /// Determines the folder, fileName, worksheetNm, range from the parameters
      /// 
      /// INPUTS:
      /// if the first parameter is a json config file then the required parameters are taken from it
      /// otherwise the required parameters are taken from the rest of the commandline.
      /// 
      /// Parameters:
      ///   arg[0]: is either a json config file or a spreadsheet path
      /// 
      /// PRECONDITIONS: none
      /// POSTCONDITIONS: 
      /// POST01: RETURNS a map of unique Excel worbook name to its configuration:
      ///   folder        
      ///   fileName      
      ///   worksheettab  
      ///   range         
      /// </summary>
      /// <param name="args"></param>
      public static Dictionary<string, Dictionary<string, string?>> 
      Init(string[] args)
      {
         Dictionary<string, Dictionary<string, string?>> map = new();

         // Parameter Validation
         if (args == null || args.Length < 1)
            throw new Exception("XL requires at least 1 argument, a json file or Excel file path, [worksheetTabNm], [range]");

         // ASSERTION: Valid cmd line parameters
         // Are the parameters supplied in a config file or the command line?
         if (Path.GetExtension(args[0]) == "xls")
            map = GetParametersFromCmdLine(args);
         else
            map = GetParametersFromCfgFile(args[0]);

         // ASSERTION: POST01
         return map;
      }

      /// <summary>
      ///   Returns a list of cfgs each cfg is a map of required parameter name to value (string, string)
      /// </summary>
      /// <param name="configFile"></param>
      /// <param name="folder"></param>
      /// <param name="fileName"></param>
      /// <param name="worksheetTabNm"></param>
      /// <param name="range"></param>
      public static Dictionary<string, Dictionary<string, string?>> 
      GetParametersFromCfgFile(string configFile)
      {
         Dictionary<string, Dictionary<string, string?>> map = new();

         var config = new ConfigurationBuilder()
           .AddJsonFile(configFile)
           .Build();

         var appSettingsSection = config.GetSection("appSettings");
         var cfgs = appSettingsSection.GetChildren();

         foreach (var cfg in cfgs)
            AddCfg(map, cfg.Key, cfg["file"], cfg["worksheet"], cfg["range"]);

         return map;
      }

      /// <summary>
      /// Gets the parameters from the commandline
      /// 
      /// Preconditions:
      ///   args contains at least 1 argument
      ///   
      /// Postconditions:
      ///   folder         = current working directory
      ///   fileName       = excel file name - mandatory 
      ///   worksheetTabNm = the worksheet tab, defaults to null meaning save the first sheet
      ///   range          = Excel range, defaults to null meaning save the entire sheet
      /// </summary>
      /// <param name="args"></param>
      /// <param name="folder"></param>
      /// <param name="fileName"></param>
      /// <param name="worksheetTabNm"></param>
      /// <param name="range"></param>
      public static Dictionary<string, Dictionary<string, string?>>
      GetParametersFromCmdLine(string[] args)
      {
         Dictionary<string, Dictionary<string, string?>> map = new();
         //var folder = Directory.GetCurrentDirectory();
         var file = args[0];
         var worksheet = args.Length > 1 ? args[1] : null;
         var range = args.Length > 2 ? args[2] : null;
         var key = Path.GetFileName(file);

         AddCfg(map, file, key, worksheet, range);
         return map;
      }

      private static void AddCfg(Dictionary<string, Dictionary<string, string?>> map
         , string key, string? file, string? worksheet, string? range)
      {
         if (worksheet == "") worksheet = null;
         if (range == "") range = null;

         map.Add(key, new Dictionary<string, string?>
            (
            new Dictionary<string, string?>()
               {
                  { "fileName"     , file},
                  { "worksheetTabNm", worksheet},
                  { "range"    , range}
               }
            ));
      }


      /// <summary>
      /// Saves all the worksheets in a workbook
      /// </summary>
      /// <param name="cfg">has 3 values:
      /// fileName
      /// range
      /// 
      /// </param>
      public static void SaveWorkbook(Dictionary<string, string?> cfg)
      {
         // Get all the worksheet tabs in the workbook and save them;
         try
         {
            // Get the workbook
            Workbook workbook    = new Workbook();
            Workbook tmpWorkbook = new Workbook();
            string? workbookName = cfg["fileName"];
            workbook.LoadFromFile(workbookName);

            // Save all worksheets using the standard range
            foreach(Worksheet worksheet in workbook.Worksheets)
               SaveWorksheet(worksheet, cfg, tmpWorkbook);
         }
         catch (Exception e)
         {
            Console.WriteLine($"500: XL.exe caught exception: {e}");
            throw;
         }
      }

      /// <summary>
      /// Saves the  given worksheet
      /// </summary>
      /// <param name="sheet"></param>
      /// <param name="map"></param>
      /// <returns></returns>
      /// <exception cref="Exception"></exception>
      public static int SaveWorksheet(Worksheet sheet, Dictionary<string, string?> map, Workbook tmpWorkbook)
      {
         if (sheet == null)
         {
            var msg = $"SaveWorksheet 030: ERROR: sheet not specified";
            Console.WriteLine(msg);
            throw new Exception(msg);
         }

         int errorCode = 0;
         string folder = Directory.GetCurrentDirectory();
         string fileName = map["fileName"] ?? "";
         string? worksheetTabNm = sheet.Name;

         string? range = map["range"];
         string xlFilePath = @$"{folder}\{fileName}";

         Console.WriteLine($"\r\nSaveWorksheet 000: starting:\r\n" +
$"filePath:       {xlFilePath}\r\n" +
$"worksheetTabNm: {worksheetTabNm}" +
$"range:          {range}");

         if (range == "") range = null;
         if (worksheetTabNm == "") worksheetTabNm = null;

         try
         {
            if (range == null)
               range = sheet.Range.RangeAddress;

            //Dictionary<string, string?> map = new Dictionary<string, string?>();

            //
            // ASSERTION: we have the sheet
            //

            // if the reference does not specify end row like A3:L
            // for example when Excel gets seriously makullit about trailing empty rows
            // then modify the range reference to range for the specified start to the real end
            int lastPopRow = sheet.LastDataRow;

            if (lastPopRow < 1)
               throw new Exception($"worksheet: {worksheetTabNm} has no rows");

            //------------------------------
            // ASSERTION: have data rows
            //------------------------------

            range = GetRange(range, lastPopRow);


            //------------------------------
            // ASSERTION: range setup
            //------------------------------

            Console.WriteLine($"SaveWorksheet 040: sheet:[{worksheetTabNm}] range:[{range}]");
            string txtFilePath = CreateTextFilePath(folder, fileName, worksheetTabNm, "txt");
            Console.WriteLine($"SaveWorksheet 050: deleting file if it exists: {txtFilePath}");

            // Initially delete the txt file if it exists
            File.Delete(txtFilePath);
            errorCode = System.Runtime.InteropServices.Marshal.GetLastWin32Error();

            // Assert the file does not exist now
            if (File.Exists(txtFilePath))
               throw new Exception($"SaveWorksheet 060: Failed to delete text file before this save [{txtFilePath}], Win GetLastError: {errorCode}");

            //----------------------------------------
            // ASSERTION: the file does not exist now
            //----------------------------------------

            // Save the worksheet as a tab separated txt file
            Console.WriteLine($"SaveWorksheet 070: Save the worksheet as a tab separated txt file");

            if (range != null)
            {
               Console.WriteLine($"SaveWorksheet 070: range specified");
               Worksheet sheet2 = tmpWorkbook.CreateEmptySheet();
               // Get the destination cell range
               CellRange srcRange = sheet.Range[range];
               CellRange tgtRange = sheet2.Range[range];

               // Copy data from the source range to the destination range
               Console.WriteLine($"SaveWorksheet 080: Copy data from the source range to the destination range");
               sheet.Copy(srcRange, tgtRange);
               Console.WriteLine($"SaveWorksheet 090: SaveToFile({txtFilePath}");
               sheet2.SaveToFile(txtFilePath, "\t", Encoding.UTF8);
               errorCode = System.Runtime.InteropServices.Marshal.GetLastWin32Error();

               // Done the copy, remove the temporary sheet
               tmpWorkbook.Worksheets.Remove(sheet2);
            }
            else
            {
               Console.WriteLine($"SaveWorksheet 100: range not specified");
               sheet.SaveToFile(txtFilePath, "\t", Encoding.UTF8);
               errorCode = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
            }

            // Postcondition: chk file exists and has been altered

            Console.WriteLine($": ");
            if (!File.Exists(txtFilePath))
               throw new Exception($"110: ERROR: Failed to save text file [{txtFilePath}], Win GetLastError: {errorCode}");

            // Display the txt file details
            var x = new System.IO.FileInfo(txtFilePath);
            Console.WriteLine($"120: Saved {Path.GetFullPath(xlFilePath)} as tsv:");
            Console.WriteLine($"130: {Path.GetFullPath(txtFilePath)} {x.Length} bytes {x.LastWriteTime}  time now: {DateTime.Now}");

            if (errorCode != 0)
               Console.WriteLine($"140: Error saving txt file ****: Win GetLastError: {errorCode}");
         }
         catch (Exception e)
         {
            Console.WriteLine($"SaveWorksheet() Error: {e}");
            throw;
         }

         return 0;
      }

      public static string? GetRange(string? range, int lastPopRow)
      {
         var parts = range?.Split(':') ?? ["", ""];
         var index = parts[1].IndexOfAny("0123456789".ToCharArray());

         if (index == -1)
         {
            var range2 = range;
            Console.WriteLine($"SaveWorksheet 035:modifying range");
            range = $"{parts[0]}:{parts[1]}{lastPopRow}";
            Console.WriteLine($"SaveWorksheet 035:modifying range {range2} to  {range}");
         }

         return range;
      }

      /// <summary>
      /// SaveWorksheet(): saves a given worksheet
      /// Called by Program.Main()
      /// Calls SaveWorksheet(Worksheet sheet, Dictionary<string, string?> map, Workbook tmpWorkbook)
      /// </summary>
      /// <param name="map">parameters for each </param>
      /// <returns></returns>
      public static int SaveWorksheet(Dictionary<string, string?> map, Workbook tmpWorkbook)
      {
         string folder = Directory.GetCurrentDirectory();
         int errorCode = 0;
         string fileName = map["fileName"] ?? "";
         string? worksheetTabNm = map["worksheetTabNm"];

         string? range = map["range"];
         string xlFilePath = @$"{folder}\{fileName}";

         Console.WriteLine($"\r\nSaveWorksheet 000: starting:\r\n" +
$"filePath:       {xlFilePath}\r\n" +
$"worksheetTabNm: {worksheetTabNm}" +
$"range:          {range}");

         if (range == "") range = null;
         if (worksheetTabNm == "") worksheetTabNm = null;

         try
         {
            // Create a Workbook instance
            Workbook workbook = new Workbook();

            // Load the Excel workbook
            Console.WriteLine($"SaveWorksheet 010: Loading workbook {xlFilePath}");
            workbook.LoadFromFile(xlFilePath);

            // Get the named worksheet, or the first worksheet if work sheet name not specified
            Worksheet sheet;
            Console.WriteLine($"SaveWorksheet 020: getting sheet");

            if (worksheetTabNm == null)
               sheet = workbook.Worksheets[0];
            else
               sheet = workbook.Worksheets[worksheetTabNm];

            errorCode= SaveWorksheet( sheet, map, tmpWorkbook);
         }
         catch(Exception e)
         {
            Console.WriteLine($"XL Error 520: caught exception {e}");
            throw;
         }

         return 0;
      }
/*
            if (range == null)
               range = sheet.Range.RangeAddress;

            if (sheet == null)
            {
               errorCode = 6;
               var msg = $"SaveWorksheet 030: ERROR: sheet {worksheetTabNm} does not exist in {xlFilePath}";
               Console.WriteLine(msg);
               throw new Exception(msg);
            }

            //
            // ASSERTION: we have the sheet
            //

            // if the reference does not specify end row like A3:L
            // for example when Excel gets seriously makullit about trailing empty rows
            // then modify the range reference to range for the specified start to the real end
            lastPopRow = sheet.LastDataRow;
            var parts = range?.Split(':') ?? ["", ""];
            var index = parts[1].IndexOfAny("0123456789".ToCharArray());

            if (index == -1)
            {
               var range2 = range;
               Console.WriteLine($"SaveWorksheet 035:modiying range");
               range = $"{parts[0]}:{parts[1]}{lastPopRow}";
               Console.WriteLine($"SaveWorksheet 035:modifying range {range2} to  {range}");
            }

            Console.WriteLine($"SaveWorksheet 040: sheet:[{worksheetTabNm}] range:[{range}]");
            //string txtFilePath = @$"{folder}\{Path.GetFileNameWithoutExtension(XlfilePath)}.txt";
            string txtFilePath = CreateTextFilePath(folder, fileName, worksheetTabNm, "txt");
            Console.WriteLine($"SaveWorksheet 050: deleting file if it exists: {txtFilePath}");
            // Initially delete the txt file if it exists
            File.Delete(txtFilePath);
            errorCode = System.Runtime.InteropServices.Marshal.GetLastWin32Error();

            // Assert the file does not exist now
            if (File.Exists(txtFilePath))
               throw new Exception($"SaveWorksheet 060: Failed to delete text file before this save [{txtFilePath}], Win GetLastError: {errorCode}");

            // Save the worksheet as a tab separated txt file
            Console.WriteLine($"SaveWorksheet 070: Save the worksheet as a tab separated txt file");

            if (range != null)
            {
               Console.WriteLine($"SaveWorksheet 070: range specified");
               Worksheet sheet2 = workbook.CreateEmptySheet();
               // Get the destination cell range
               CellRange srcRange = sheet.Range[range];
               CellRange tgtRange = sheet2.Range[range];

               // Copy data from the source range to the destination range
               Console.WriteLine($"SaveWorksheet 080: Copy data from the source range to the destination range");
               sheet.Copy(srcRange, tgtRange);
               Console.WriteLine($"SaveWorksheet 090: SaveToFile({txtFilePath}");
               sheet2.SaveToFile(txtFilePath, "\t", Encoding.UTF8);
               errorCode = System.Runtime.InteropServices.Marshal.GetLastWin32Error();

               // Done the copy, remove the temporary sheet
               workbook.Worksheets.Remove(sheet2);
            }
            else
            {
               Console.WriteLine($"SaveWorksheet 100: range not specified");
               sheet.SaveToFile(txtFilePath, "\t", Encoding.UTF8);
               errorCode = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
            }

            // Postcondition: chk file exists and has been altered

            Console.WriteLine($": ");
            if (!File.Exists(txtFilePath))
               throw new Exception($"110: ERROR: Failed to save text file [{txtFilePath}], Win GetLastError: {errorCode}");

            // Display the txt file details
            var x = new System.IO.FileInfo(txtFilePath);
            Console.WriteLine($"120: Saved {Path.GetFullPath(xlFilePath)} as tsv:");
            Console.WriteLine($"130: {Path.GetFullPath(txtFilePath)} {x.Length} bytes {x.LastWriteTime}  time now: {DateTime.Now}");

            if (errorCode != 0)
               Console.WriteLine($"140: Error saving txt file ****: Win GetLastError: {errorCode}");
         }
         catch (Exception e)
         {
            Console.WriteLine($"500: XL.exe caught exception: {e}");
            throw;
         }

         return errorCode;
      }
*/

      public static string CreateTextFilePath(string? folder, string? workBookName, string? workSheetName, string extension)
      {
         if (workBookName == null)
            throw new Exception($"CreateTextFilePath() failed: null workBookName");

         if (folder == null)
            throw new Exception($"CreateTextFilePath(workBookName: {workBookName}) failed: null folder");

         if (workSheetName == null)
            throw new Exception($"CreateTextFilePath(workBookName: {workBookName}) failed: null workSheetName");

         string relFolder = Path.GetDirectoryName(workBookName) ?? "";

         string fileNameNoExt = Path.GetFileNameWithoutExtension(workBookName);
         string folder2 = relFolder == "" ? @$"{folder}\{fileNameNoExt}.{workSheetName}.{extension}"
                                         : @$"{folder}\{relFolder}\{fileNameNoExt}.{workSheetName}.{extension}";
         return folder2;
      }
   }
}
