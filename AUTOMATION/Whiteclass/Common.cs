using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
//using White.Support_Files;
using System.Windows.Forms;
using System.Collections;
using TestStack.White.InputDevices;
using TestStack.White.WindowsAPI;
using Microsoft.Office.Interop.Excel;


namespace Whiteclass
{
    public static class Common

    {
        public static string sTestCaseName = "";
        public static string sTestCaseStatus = "";
        public static char identifier;
        public static Hashtable ObjectIdentifier = new Hashtable();
        public static string sMachineName = Environment.MachineName;
        // public static string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

       
        // public static char identifierList[]
        public struct st_DebugConfig
        {
            public string sLocalPath;
            public string sDebugConfigFileName;
            public bool bCopyLocators;
            public string sTID;
            public string sTestQueueTable;
            // Variable to decide whether to overwrite the temp files
            public bool bOverwrite;
            public string sIsScheduledTest;
          //  public string sFailureEmail;
          //  public string sSuccessEmail;
            public string sTestCaseID;
            public string sTQDID;
            public string sResultsFile;
            public string sResultsFileLocation;
            public string sTFSTestCaseID;
            public string sunionTestCases;
            public string sMode;

            public int iUnionTestCount;
            public int iCurrentTestCount;
            public string sUnionTestStatus;

           


        }

        /// <summary>
        ///    //Global Test Support Variables
        /// </summary>
        public struct st_GlobalTestSupport
        {
            public string sTestCaseName;
            public bool bTELImplemented;
            public bool bResult;
            public int iwaitTime; //2 Sec
            public int iTimeDelay; //1 Min
            public int iTime; //5 Sec
            public int iWhileTime; //5 Sec
            public int iTabTime; //5 Sec
            public int iElapsedTime; // 1 Min
            public int iIdleTime; //2 mins
            public int iInActiveTime; // 5 Min's
            public List<string> lFormatResult;
            public string sLocatorValue;
        }

        //Configuration Variables
        public struct st_Config
        {
            //My prj Launch URL 

            public string sMyPrjPath;
            public string sMyPrjBuild;
            public string sMyPrjCopyrightYear;

            public string sMyPrjUserName;
            public string sMyPrjPassword;


            public string sLocatorFile;
            public string sLocatorSheet;
            public string sLanguageTag;
            public string sCsvFilepath;
            public string sTareFilePath;
            public string sImportPath;
            public string sExportPath;
            public string sInstallPath;


            public string sPrjType;
            public string sDomain;
            public string sResultfile;

            public string sEnvironment;
        }

          
        public static st_DebugConfig obj_st_DebugConfig = new st_DebugConfig();
        public static st_Config obj_st_Config = new st_Config();
        //public static st_Location obj_st_Location = new st_Location();
        public static st_GlobalTestSupport obj_st_GlobalTestSupport = new st_GlobalTestSupport();
        public static List<string> lslines = new List<string>();


        
       public static void Sleep(int TimeInSec)
        {
            //Thread.Sleep(TimeSpan.FromSeconds(TimeInSec));
            Thread.Sleep((int)TimeSpan.FromSeconds(TimeInSec).TotalMilliseconds);

        }

        public static void CloseAll(Boolean bLogStep)
        {
            if (Clipboard.ContainsText())
            {
                Clipboard.Clear();
            }
            if (bLogStep)
            {
                //ATEL.LogTestSteps("Closing AUT and all windows", -1);

            }
            try
            {
                fnKillWinProcesses();
                Sleep(10);
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Closing AUT and all windows" + e.Message);
            }
        }

        public static void fnKillWinProcesses()
        {
            List<string> ProcessToKill = new List<string>() { "TPWICP"  };//"dllhost",  "msl","IEXPLORE", "EXCEL", "MaxExcelAndClose" 
            foreach (string sPro in ProcessToKill)
            {
                foreach (Process proc in Process.GetProcessesByName(sPro))
                    proc.Kill();
            }
        }

        public static void fnKillWinProcesses(string sProcess)
        {
            foreach (Process proc in Process.GetProcessesByName(sProcess))
            {
                proc.Kill();
            }
        }

        
        /// <summary>
        /// Get All the Configuration from Configuration files and assign to Global Variables
        /// </summary>
        public static void GetAutoConfigData()
        {
                       obj_st_DebugConfig.sLocalPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;// @"C:\QA10\SOTAXQAAUTOMATION";
            obj_st_DebugConfig.sDebugConfigFileName = "TestAutoConfig.ini";
             obj_st_DebugConfig.bCopyLocators = true;
            obj_st_DebugConfig.bOverwrite = true;
          



            // string path = Directory.GetCurrentDirectory();

             string sConfigPath = obj_st_DebugConfig.sLocalPath +"\\"+ obj_st_DebugConfig.sDebugConfigFileName;

           // string sConfigPath =   @"C:\QA10\SOTAXQAAUTOMATION\TestAutoConfig.ini";

            try
            {
                if (System.IO.File.Exists(sConfigPath) == true)
                {
                    IniFile ini = new IniFile(sConfigPath);

                    //Read all Config Structure Variables

                    

                    // ******** Intl Location *************
                    obj_st_Config.sLocatorFile = ini.Read("FILE", "INTL");
                    obj_st_Config.sLocatorSheet = ini.Read("LOCATORSHEET", "INTL");
                    obj_st_Config.sLanguageTag = ini.Read("LANGUAGE", "INTL");
                    

                    


                    // ******** My Prj Location *************
                    obj_st_Config.sMyPrjPath = ini.Read("Path", "MYPRJ");
                    obj_st_Config.sMyPrjBuild = ini.Read("Build", "MYPRJ").ToLower();
                    obj_st_Config.sMyPrjCopyrightYear = ini.Read("CopyrightYear", "MYPRJ").ToLower();
                    obj_st_Config.sMyPrjUserName = ini.Read("UserName", "MYPRJ");
                    obj_st_Config.sMyPrjPassword = ini.Read("Password", "MYPRJ");

                    obj_st_Config.sPrjType = ini.Read("PrjType", "MYPRJ");
                    obj_st_Config.sDomain = ini.Read("Domain", "MYPRJ") ;
                    obj_st_Config.sResultfile = ini.Read("Resultfile", "MYPRJ");
                    obj_st_Config.sTareFilePath = ini.Read("Tarefile", "MYPRJ");
                    obj_st_Config.sImportPath = ini.Read("ImportPath", "MYPRJ");
                    obj_st_Config.sExportPath = ini.Read("ExportPath", "MYPRJ");
                    obj_st_Config.sInstallPath = ini.Read("InstallPath", "MYPRJ");
                    if (obj_st_Config.sLocatorSheet.Equals("MYPRJ"))
                    {
                        if (obj_st_Config.sMyPrjPath == "")
                            throw new Exception("Could not find the JAA Path in the config file");
                    }

                    obj_st_Config.sEnvironment = ini.Read("Environment", "ENVIRONMENT");
                    if (obj_st_Config.sEnvironment == "")
                        throw new Exception("Could not find the Environment in the config file");
                    //Read all Debug Config Strcuture Variables
                    if (ini.Read("COPYTAGS", "Tag").ToLower() == "true")
                        obj_st_DebugConfig.bCopyLocators = true;
                    else if (ini.Read("COPYTAGS", "Tag").ToLower() == "false")
                        obj_st_DebugConfig.bCopyLocators = false;
                    else
                        throw new Exception("Could not find the Copy Tags boolean in the debug config file");

                    if (ini.Read("OVERWRITE", "TEMPFILES").ToLower() == "true")
                        obj_st_DebugConfig.bOverwrite = true;
                    else if (ini.Read("OVERWRITE", "TEMPFILES").ToLower() == "false")
                        obj_st_DebugConfig.bOverwrite = false;
                    else
                        throw new Exception("Could not find the overwrite temp files boolean in the debug config file");

                }
                else
                {
                    throw new FileNotFoundException("File " + obj_st_DebugConfig.sDebugConfigFileName + " not found in the path " + obj_st_DebugConfig.sLocalPath);
                }
            }
            catch (Exception e)
            {
                
                Assert.Fail(e.Message);
                throw new FileNotFoundException("File " + obj_st_DebugConfig.sDebugConfigFileName + " not found in the path " + obj_st_DebugConfig.sLocalPath);
            }
        }

        public class IniFile
        {
            string Path;
            string EXE = Assembly.GetExecutingAssembly().GetName().Name;

            [DllImport("kernel32")]
            static extern long WritePrivateProfileString(string Section, string Key, string Value, string FilePath);

            [DllImport("kernel32")]
            static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);

            public IniFile(string IniPath = null)
            {
                Path = new FileInfo(IniPath ?? EXE + ".ini").FullName.ToString();
            }

            public string Read(string Key, string Section)
            {
                var RetVal = new StringBuilder(255);
                GetPrivateProfileString(Section ?? EXE, Key, "", RetVal, 255, Path);
                return RetVal.ToString();
            }

            public void Write(string Key, string Value, string Section)
            {
                WritePrivateProfileString(Section ?? EXE, Key, Value, Path);
            }

            public void DeleteKey(string Key, string Section = null)
            {
                Write(Key, null, Section ?? EXE);
            }

            public void DeleteSection(string Section = null)
            {
                Write(null, null, Section ?? EXE);
            }

            public bool KeyExists(string Key, string Section = null)
            {
                return Read(Key, Section).Length > 0;
            }

        }

        // *************Locator Methods***********************************************************************************

        public static System.Data.DataTable dDataTable = null;
        public static DataSet dsTemp = null;

       public static void LoadLocatorToDataTable()
        {
            try
            {
                obj_st_GlobalTestSupport.bTELImplemented = true;
                dsTemp = new DataSet();
                System.Data.OleDb.OleDbDataAdapter MyCommand;

               // string sConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + Common.obj_st_Config.sLocatorFile + ";" + "Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";

               string sConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source = '" + Common.obj_st_Config.sLocatorFile + "';Extended Properties=\"Excel 8.0;HDR=YES;\"";
                string sql = "select Elements," + Common.obj_st_Config.sLanguageTag + " from [" + Common.obj_st_Config.sLocatorSheet + "$]";

                System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection(sConn);
                MyConnection.Open();
                MyCommand = new System.Data.OleDb.OleDbDataAdapter(sql, MyConnection);
                MyCommand.TableMappings.Add("Table", "LocatorDataTable");
                MyCommand.Fill(dsTemp, "LocatorDataTable");
                dDataTable = dsTemp.Tables["LocatorDataTable"];
                MyConnection.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception is : " + ex);
                throw new Exception("Unable to Load Locator Sheet To DataTable : " + ex);
            }
        }

       
        public static string sElementLogName;

        

        /// <summary>
        ///  Global Test Support Variables :  iwaitTime:2 Sec, iTimeDelay: 1Min, iSleepTime:5 Sec, iTabTime:30 Sec,iElapsedTime:3Min , iInActiveTime:10 Min's
        /// </summary>
        public static void SetAllGlobalValues()
        {
            //Function to Set all the global values like timers etc
            obj_st_GlobalTestSupport.iwaitTime = 2;
            obj_st_GlobalTestSupport.iTimeDelay = 10;
            obj_st_GlobalTestSupport.iTime = 5;
            obj_st_GlobalTestSupport.iWhileTime = 7;
            obj_st_GlobalTestSupport.iTabTime =50 ;
            obj_st_GlobalTestSupport.iElapsedTime = 180;
            obj_st_GlobalTestSupport.iIdleTime = 300;
            obj_st_GlobalTestSupport.iInActiveTime = 600;
            
        }


        public static string SQLSafeQuotes(string sStr)
        {

            if (sStr.Contains("'"))
            {
                sStr = sStr.Replace("'", "\"");
            }
            return sStr;
        }


        
        public static void Exception(Exception e, Boolean bScreenshot = true)
        {

            Common.sTestCaseStatus = "FAIL";
            Console.WriteLine("ERROR:" + e.Message);
            
            Assert.Fail(e.Message);
        }

        public static RegistryKey Regkey(String RegistryType, String RegistryLocation)
        {
            RegistryKey key;
            switch (RegistryType)
            {
                case "LocalMachine":
                    key = Registry.LocalMachine.OpenSubKey(@RegistryLocation, true);
                    break;
                case "CurrentUser":
                    key = Registry.CurrentUser.OpenSubKey(@RegistryLocation, true);
                    break;
                default:
                    key = Registry.LocalMachine.OpenSubKey(@RegistryLocation, true);
                    break;
            }
            return key;
        }
        public static string getRegistryKeyValue(String RegistryType, String RegistryLocation, String ValueName)
        {
            RegistryKey key = Regkey(RegistryType, RegistryLocation);
            string data = key.GetValue(ValueName).ToString();
            //ATEL.LogTestSteps("Registry Key Data: " + data + "</font>");
            key.Close();
            return data;
        }

        public static void setRegistryKeyValue(String RegistryType, String RegistryLocation, String ValueName, String value)
        {
            RegistryKey key = Regkey(RegistryType, RegistryLocation);
            //adding/editing a value 
            key.SetValue(ValueName, value); //sets 'someData' in 'someValue' 
            key.Close();
        }

        public static void deleteRegistry(String RegistryType, String RegistryLocation, String ValueName)
        {
            try
            {
                RegistryKey key = Regkey(RegistryType, RegistryLocation);
                if (key != null)
                {
                    if (key.GetValue(ValueName).ToString() != null)
                        key.DeleteValue(ValueName); //delete if exist
                    else
                        Console.WriteLine("Value not found");
                }
                else
                {
                    Console.WriteLine("key not found");
                }
            }
            catch
            {
                //Not hangeling exception
            }

        }

        //*Function to change Language in Regestry based on Selected Language *//
        public static void fnSetLanguageInRegistry()
        {
            switch (Common.obj_st_Config.sLanguageTag.ToUpper())
            {
                case "ENGLISH":
                    
                        setRegistryKeyValue("LocalMachine", "SOFTWARE\\WOW6432Node\\Packard\\Multiprobe II", "Culture", "");
                    break;
                
                default:
                         setRegistryKeyValue("LocalMachine", "SOFTWARE\\WOW6432Node\\Packard\\Multiprobe II", "Culture", "");
                    break;
            }
        }

        
        public static void InitializeSettings()
        {
            obj_st_DebugConfig.sLocalPath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;// @"C:\QA10\SOTAXQAAUTOMATION\";



            obj_st_DebugConfig.sDebugConfigFileName = obj_st_DebugConfig.sLocalPath + "\\TestAutoConfig.ini"; //@"C:\QA10\SOTAXQAAUTOMATION\TestAutoConfig.ini";

            GetAutoConfigData();
            
            SetAllGlobalValues();
            LoadLocatorToDataTable();
            
            //function kill process
            fnKillWinProcesses();
            Sleep(10);
            

        }

        public static void CleanupSettings()
        {
            Keyboard.Instance.HoldKey(KeyboardInput.SpecialKeys.RETURN);
            Keyboard.Instance.LeaveKey(KeyboardInput.SpecialKeys.RETURN);
            CloseAll(false);
        }

        


        /// <summary>
        /// Get substring from Config File
        /// </summary>
        /// <param name="sFileName">File Name</param>
        /// <param name="sString">Search Contains value</param>
        /// <param name="sSplitChar">Char to split</param>
        /// <param name="sLogTestStep">Display log</param>
        /// <returns></returns>
        public static string GetConfigValue(string sFileName, string sString, string sLogTestStep)
        {
            string sGetString = "";
            try
            {
                using (StreamReader sr = new StreamReader(sFileName))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.Length != 0 & line.Contains(sString))
                        {
                            sGetString = line.Split('>')[1];
                            sGetString = sGetString.Split('<')[0];
                            break;
                        }
                    }
                }
                // ATEL.LogTestSteps(sLogTestStep + " " + sGetString.Trim());
            }
            catch (Exception e)
            {
                Exception(e);
            }
            return sGetString.Trim();

        }

        /// <summary>
        /// Get a particular value from any csv files
        /// </summary>
        /// <param name="sFileName">File Name</param>
        /// <param name="sSheetName">Sheet name</param>
        /// <param name="iRow">Row to fetch the value</param>
        /// <param name="iColumn">column to fetch the value</param>
        /// <returns = "sGetString">fetched value from csv file</returns>
        public static string fetchValueFromCSVFile(string sFileName, string sSheetName, int iRow, int iColumn)
        {
            string sGetString = "";
            try
            {
                FileInfo file = new FileInfo(sFileName);

                string sConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + file.DirectoryName + "; Extended Properties = 'text;HDR=Yes;FMT=Delimited(,)'; ";
                using (System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection(sConn))
                {
                    using (System.Data.OleDb.OleDbCommand MyCommand = new System.Data.OleDb.OleDbCommand(string.Format
                                              ("SELECT * FROM [{0}]", file.Name), MyConnection))
                    {
                        MyConnection.Open();

                        // Using a DataTable to process the data
                        using (System.Data.OleDb.OleDbDataAdapter adp = new System.Data.OleDb.OleDbDataAdapter(MyCommand))
                        {

                            System.Data.DataTable  dDataTable = new System.Data.DataTable("MyTable");
                            adp.Fill(dDataTable);

                            sGetString = dDataTable.Rows[iRow][iColumn].ToString();

                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception is : " + ex);
                throw new Exception("Unable to Load Locator Sheet To DataTable : " + ex);
            }
            return sGetString;
        }


        /// <summary>
        /// Get the column header from any csv files
        /// </summary>
        /// <param name="sFileName">File Name</param>
        /// <param name="sSheetName">Sheet name</param>
        /// <param name="iColumn">column to fetch the header</param>
        /// <returns = "sGetString">fetched header from csv file</returns>
        public static string getCSVColumnHeader(string sFileName, string sSheetName, int iColumn)
        {
            string sGetString = "";
            try
            {
                FileInfo file = new FileInfo(sFileName);

                string sConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + file.DirectoryName + "; Extended Properties = 'text;HDR=Yes;FMT=Delimited(,)'; ";
                using (System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection(sConn))
                {
                    using (System.Data.OleDb.OleDbCommand MyCommand = new System.Data.OleDb.OleDbCommand(string.Format
                                              ("SELECT * FROM [{0}]", file.Name), MyConnection))
                    {
                        MyConnection.Open();

                        // Using a DataTable to process the data
                        using (System.Data.OleDb.OleDbDataAdapter adp = new System.Data.OleDb.OleDbDataAdapter(MyCommand))
                        {
                            // DataTable dDataTable = new DataTable("MyTable");
                            System.Data.DataTable dDataTable = new System.Data.DataTable("MyTable");
                            adp.Fill(dDataTable);
                            //dDataTable.Columns

                            sGetString = dDataTable.Columns[1].ToString();

                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception is : " + ex);
                throw new Exception("Unable to Load Locator Sheet To DataTable : " + ex);
            }
            return sGetString;
        }

        

        public static String FindPageElement(string sPageObject)
        {
            obj_st_GlobalTestSupport.sLocatorValue = GetLocator(sPageObject).Remove(0, 1);
            Common.sElementLogName = sPageObject.Substring(sPageObject.LastIndexOf('.') + 4);

            string sObject = (string)GetLocator(sPageObject);
            string propertyValue = sObject;

            if (!sPageObject[0].Equals('_'))
            {
                identifier = sObject[0];
                propertyValue = sObject.Substring(1, sObject.Length - 1);

                ObjectIdentifier[propertyValue] = identifier;

                // ATEL.LogTestSteps("Display Identifier for Debugging purpose:" + identifier);
            }
            return propertyValue;
        }
        private static string GetLocator(string sObject)
        {

            string sLocator = null;
            string sDataFilterExp = "Elements='" + sObject + "'";
            
            string sENV = Common.obj_st_Config.sEnvironment;
            string[] property = null;
            try
            {

                sLocator = dDataTable.Select(sDataFilterExp)[0][1].ToString();
                if (sLocator == "")
                {
                    throw new Exception("ERROR: " + sObject + "  missing in Locator sheet. ");
                }
                Console.WriteLine("In data table ************  Selecting  " + sLocator);
                property = sLocator.Split(new Char[] { '|' });
                switch (sENV.ToLower())
                {
                    case "qa":
                        Console.WriteLine("In data table ************ QA  Selecting  " + sLocator);
                        sLocator = property[0];
                        break;
                    case "ruo":
                        Console.WriteLine("In data table ************ RUO  Selecting  " + sLocator);
                        sLocator = property[0];
                        break;
                    case "production":
                        Console.WriteLine("In data table ************ PROD Selecting  " + sLocator);
                        if (property.GetLength(0) <= 1)
                        {
                            sLocator = property[0];
                        }
                        else sLocator = property[2];
                        break;
                    default:
                        Console.WriteLine("In data table ************ Default Selecting  " + sLocator);
                        sLocator = property[0];
                        break;
                }
            }
            catch (IndexOutOfRangeException ex)
            {
                throw new Exception("ERROR: " + sObject + "  missing in Locator sheet. " + ex);
            }
            return sLocator;
        }
        /// <summary>
        /// save the test result status to excel file
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="tcname"></param>
        /// <param name="status"></param>
        public static void SaveResultsToExcel(string sheetName, int row, int col, string tcname, string status)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application oXL = null;
                Microsoft.Office.Interop.Excel._Workbook oWB = null;
                Microsoft.Office.Interop.Excel._Worksheet oSheet = null;

                try
                {
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oWB = oXL.Workbooks.Open(@"C:\QA10\SOTAXQAAUTOMATION\Results.xls", ReadOnly: false);

                    oXL.Visible = false;
                    oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];

                    oSheet.Cells[row, col] = tcname;
                    if (status == "PASS")
                        oSheet.Cells[row, col + 1] = "PASS";
                    else
                        oSheet.Cells[row, col + 1] = "FAIL";
                    oSheet.Cells[row, col + 2] = DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss");



                    oWB.Save();

                    //   oWB.Close();


                    //  MessageBox.Show("Done!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    if (oWB != null)
                        oWB.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception is : " + ex);
                throw new Exception("Unable to Load Locator Sheet To DataTable : " + ex);
            }
        }

        public static void SaveResultstepToExcel(string sheetName, string ActionName, string Testname, string Results, string status)
        
        {

            System.Data.DataTable rowdt = null;

            dsTemp = new DataSet();

            // string filename = @"C:\QA10\SOTAXQAAUTOMATION\Results.xls";

            string filename = obj_st_Config.sResultfile; //once prjtype is ppases comment line


            System.Data.OleDb.OleDbDataAdapter MyCommand;

         

            string sConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source = '" + filename + "';Extended Properties=\"Excel 8.0;HDR=YES;\"";
            System.Data.OleDb.OleDbConnection MyConnection = new System.Data.OleDb.OleDbConnection(sConn);
            MyConnection.Open();
            string sql = "Select Count(SlNo) from [" + sheetName + "$]";

            MyCommand = new System.Data.OleDb.OleDbDataAdapter(sql, MyConnection);
            MyCommand.TableMappings.Add("Table", "Rowcount");
            MyCommand.Fill(dsTemp, "Rowcount");
            rowdt = dsTemp.Tables["Rowcount"];

            int i = Convert.ToInt32(rowdt.Rows[0][0]);

            if (i == 0)
                i = 1;
            else
                i = i + 1;
            rowdt.Clear();

            if (status == "TRUE")
                status = "PASS";
            else
                status = "FAIL";


            string InsertSql = "Insert into [" + sheetName + "$](SlNo,ActionName,Testname,Results,ExecutionDate,ExecutionStatus)Values('" + i + "','" + ActionName + "','" + Testname + "','" + Results + "','" + DateTime.Now.ToString() + "','" + status + "')";

            System.Data.OleDb.OleDbCommand InCommand = new System.Data.OleDb.OleDbCommand();
            InCommand.CommandText = InsertSql;
            InCommand.Connection = MyConnection;
            InCommand.ExecuteNonQuery();
            rowdt.Clear();
            MyConnection.Close();
            

        }

        public static void ClearExcel(string sheetName)
        {
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;
            // string sConfigPath = @"D:\TPW\QA\SOTAXQAAUTOMATION\TestAutoConfig.ini";
            // IniFile ini = new IniFile(sConfigPath);

            // oWB = oXL.Workbooks.Open(ini.Read("FILE", "RESULTSLOC"), ReadOnly: false);

            oXL = new Microsoft.Office.Interop.Excel.Application();
            //  oWB = oXL.Workbooks.Open(@"C:\QA10\SOTAXQAAUTOMATION\Results.xls", ReadOnly: false);
            oWB = oXL.Workbooks.Open(obj_st_Config.sResultfile, ReadOnly: false);

            // myExcelApp.Workbooks.Open(Filename: @"F:\oo\bar.xlsx", ReadOnly: false);
            oXL.Visible = false;
            oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];
            int rowcnt = oSheet.UsedRange.Rows.Count;

            {
                for (int i = rowcnt; i >= 2; i += -1)
                    // oSheet.Rows(i).Delete();
                    ((Range)oSheet.Rows[i]).Delete();
            }

            oWB.Save();
            oWB.Close();

        }


    }
}

