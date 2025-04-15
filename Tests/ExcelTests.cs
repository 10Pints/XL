using Microsoft.VisualStudio.TestTools.UnitTesting;
using static XL_Lib.Excel;
using XL;

namespace XL_Lib.Tests
{
   [TestClass()]
   public class ExcelTests
   {
      /// <summary>
      ///  folder: 
      ///  "file": "Class Schedule.xlsx",
      ///  "worksheet": "Schedule 250217",
      ///  "range": "A2:P"
      /// </summary>
      [TestMethod]
      public void LaunchUsingJsonCfgFileTest()
      {
         var cfgfileName = "SaveDorsuSpreadsheetsAsTxt_config.json";
         string[] args = [cfgfileName];
         Dictionary<string, Dictionary<string, string?>> map = GetParametersFromCmdLine(args);

         Assert.AreNotEqual(null, map);
         Assert.AreEqual(1, map.Count);

         Assert.AreEqual(3, map.First().Value.Count);
         var ret = Program.Main(args);
         Assert.AreEqual(0, ret, "expected success but an error occurred");

         // Check the file
         string folder = Directory.GetCurrentDirectory();
         Dictionary<string, Dictionary<string, string?>> cfg = Excel.Init(args);
         Dictionary<string, string?> cfg1 = cfg.First().Value;
         var xlFileName = cfg1["fileName"];
         var worksheetTabNm = cfg1["worksheetTabNm"];
         var range= cfg1["range"];
         string txtFilePath = CreateTextFilePath(folder, xlFileName, worksheetTabNm, "txt");
         // Read the file
         var lines = File.ReadAllLines(txtFilePath);
         // Chk the hdr row has tab separators
         var line = lines[0];
         Assert.IsTrue(line.Contains('\t'));
         line = lines[1];
         Assert.IsTrue(line.Contains('\t'));
      }

      [TestMethod]
      public void GetParametersFromCmdLineTest()
      {
         var fileName = "args.xls";
         var worksheetTabNm = "tab1";
         var range = "A5:R";
         string[] args = [fileName, worksheetTabNm, range];
         Dictionary<string, Dictionary<string, string?>> map = GetParametersFromCmdLine(args);

         Assert.AreNotEqual(null, map);
         Assert.AreEqual(1, map.Count);
         var cfg = map.First().Value;

         Assert.AreEqual(3, cfg.Count);
         Assert.AreEqual(fileName, cfg["fileName"]);
         Assert.AreEqual(worksheetTabNm, cfg["worksheetTabNm"]);
         Assert.AreEqual(range, cfg["range"]);
      }

      [TestMethod()]
      public void GetParametersFromCfgFileTest()
      {
         string configFile = "GetParametersFromCfgFileTest.json";
         Dictionary<string, Dictionary<string, string?>> cfg_map = GetParametersFromCfgFile(configFile);
         Assert.AreNotEqual(null, cfg_map);
         Assert.AreEqual(4, cfg_map.Count);
         ChkCfgHelper(cfg_map, "AllSheets", 3, "GMeet Attendance Report.xlsx", "*", "A1:G");
         ChkCfgHelper(cfg_map, "Pathogens ImportCorrections", 3, "ImportCorrections_221018-Pathogens_A-C.xlsx", "Pathogens", "A1:R5000");
      }

      [TestMethod()]
      public void SaveAllsheetsTest()
      {
         string configFile = "SaveAllsheetsTest.json";
         Dictionary<string, Dictionary<string, string?>> cfg_map = GetParametersFromCfgFile(configFile);
         Assert.AreNotEqual(null, cfg_map);
         Assert.AreEqual(1, cfg_map.Count);
         ChkCfgHelper(cfg_map, "AllSheets", 3, ".\\Attendance\\GMeet Attendance Report.xlsx", "*", "A1:G");

         Program.Main(new string[]{ configFile });
      }

      protected void ChkCfgHelper(
         Dictionary<string, Dictionary<string, string?>> cfg_map
         , string key
         , int exp_cnt
         , string? exp_fileName
         , string? exp_worksheetTabNm
         , string? exp_range
         )
      {
         Dictionary<string, string?> cfg = cfg_map[key];
         Assert.AreEqual(exp_cnt, cfg.Count);
         Assert.AreEqual(exp_fileName, cfg["fileName"]);
         Assert.AreEqual(exp_worksheetTabNm, cfg["worksheetTabNm"]);
         Assert.AreEqual(exp_range, cfg["range"]);
      }

      [TestMethod()]
      public void InitTest()
      {
         string[] args = ["GetParametersFromCfgFileTest.json"];
         Dictionary<string, Dictionary<string, string?>> exp = new Dictionary<string, Dictionary<string, string?>>()
         {
            {
            "AllSheets",
                 new  Dictionary<string, string?>()
                 {
                    { "fileName" , "GMeet Attendance Report.xlsx"},
                    { "worksheetTabNm", "*"},
                    { "range"    , "A1:G"}
                 }
            },
            {
            "Company ImportCorrections", 
                  new  Dictionary<string, string?>()
                  {
                     { "fileName", "ImportCorrections_221018-Company.xlsx"},
                     { "worksheetTabNm", null},
                     { "range", null}
                  }
            },
            {
               "Crops ImportCorrections",
                  new  Dictionary<string, string?>()
                  {
                     { "fileName", "ImportCorrections_221018-Crops.xlsx"},
                     { "worksheetTabNm", null},
                     { "range", null}
                  }
            }
            ,{
               "Pathogens ImportCorrections",
                  new  Dictionary<string, string?>()
                  {
                     { "fileName", "ImportCorrections_221018-Pathogens_A-C.xlsx"},
                     { "worksheetTabNm", "Pathogens"},
                     { "range", "A1:R5000"}
                  }
            }

         };

         InitTestHelper(args, exp);
      }

      protected void InitTestHelper(string[] args, Dictionary<string, Dictionary<string, string?>> exp_cfgs)
      {
         Dictionary<string, Dictionary<string, string?>> act_cfgs = Init(args);

         Assert.AreEqual(exp_cfgs.Count, act_cfgs.Count);

         foreach(KeyValuePair<string, Dictionary<string, string?>> exp_cfg_kvp in exp_cfgs)
         { 
            var key = exp_cfg_kvp.Key;
            var exp_cfg = exp_cfg_kvp.Value;
            var act_cfg = act_cfgs[key];

            foreach(var exp_kvp in exp_cfg)
            {
               var key2= exp_kvp.Key;
               var exp_value = exp_kvp.Value;
               var act_value = act_cfg[key2];
               Assert.AreEqual(exp_value, act_value);
            }
         }
      }
   }
}

/*
Expected:<mportCorrections_221018-Pathogens_A-C.xlsx>. 
Actual  :<ImportCorrections_221018-Pathogens_A-C.xlsx>. 
 */