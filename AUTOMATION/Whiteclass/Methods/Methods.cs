using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Castle.Core;
using TestStack.White;
using TestStack.White.Utility;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.InputDevices;
using TestStack.White.WindowsAPI;
using TestStack.White.UIItems.TreeItems;
using TestStack.White.UIItems.WindowStripControls;
//using Whiteclass.Common;
using TestStack.White.Factory;
using TestStack.White.UIItems;
using System.Windows.Automation;

namespace Whiteclass.Methods
{
    class Methods
    {
        public static TestStack.White.Application application;
        public static Window formWindow;


        internal static object GetWindows()
        {
            throw new NotImplementedException();
        }

     public static void AppstLaunchSample(Boolean bLogin = true)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = Whiteclass.Common.obj_st_Config.sMyPrjPath;
                startInfo.WorkingDirectory = Path.GetDirectoryName(startInfo.FileName);               

                Process.Start(startInfo);
                application = Application.AttachOrLaunch(startInfo);
                Common.ClearExcel("Action");
                formWindow = Retry.For(() => application.GetWindows().First(x => x.Title.Contains(pgHomeWindow._HomeWindow())), TimeSpan.FromSeconds(30));
                CommonMethods.setText(formWindow, pgHomeWindow.txtname(), pgHomeWindow.txtname(), "123");

                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                if (Methods.VerifyFieldCheck(formWindow) == true)
                {
                    Common.SaveResultstepToExcel("Action", "LaunchForm", "Check whether name can have integers  eg 123", "Err msg displayed ", "TRUE");

                    CommonMethods.setText(formWindow, pgHomeWindow.txtname(), pgHomeWindow.txtname(), "jesnyantony");

                    Common.SaveResultstepToExcel("Action", "LaunchForm", "Check whether name can have alphabets  eg jesnyantony", "sucess ", "TRUE");
                 }
                CommonMethods.setText(formWindow, pgHomeWindow.txtpass(), pgHomeWindow.txtpass(), "xxx");

                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                CommonMethods.ClickButton(Methods.formWindow, pgHomeWindow.btn_Submit());
                Common.SaveResultstepToExcel("Action", "LaunchForm", "Check whether invalid password eg xxx", "Err msg displayed ", "TRUE");
                TimeSpan.FromSeconds(50);

                CommonMethods.setText(formWindow, pgHomeWindow.txtpass(), pgHomeWindow.txtpass(), "jesny");

                Common.SaveResultstepToExcel("Action", "LaunchForm", "Check whether valid password result eg.jesny", "sucess ", "TRUE");
                
                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                formWindow.WaitWhileBusy();
                CommonMethods.ClickButton(Methods.formWindow, pgHomeWindow.btn_Submit());
                //string str =CommonMethods.getLabelText(formWindow, pgHomeWindow.lbl_Output());
                formWindow.WaitWhileBusy();
                CommonMethods.ClickButton(Methods.formWindow, pgHomeWindow.btn_Close());
                Common.SaveResultstepToExcel("Action", "LaunchForm", "Form2 opened and closed", "success ", "TRUE");
                CommonMethods.ClickButton(Methods.formWindow, pgHomeWindow.btnCancel());
                Common.SaveResultstepToExcel("Action", "LaunchForm", "Form1 closed", "success ", "TRUE");
                TimeSpan.FromSeconds(50);
                //formWindow.Close();
            }

            catch (Exception e)
            {
                Common.Exception(e);
            }
        }


            
        internal static void Click(string strMethod)
        {
            throw new NotImplementedException();
        }
    
    public static string FormatSotaxTitle(String sTitle, string sUserName = null, string sProjectType = "")
    {
        string sSotaxFormattedTitle = "";
        try
        {
            //if (sTitle != " ")
            //{
            //    sSotaxFormattedTitle = sTitle;
            //}
            //else
            //{
                if (Common.obj_st_Config.sPrjType.ToUpper().Equals("TPW"))
                    sSotaxFormattedTitle = sTitle.Replace("{0}", "TPW");
                else if(Common.obj_st_Config.sPrjType.ToUpper().Equals("APW"))
                    sSotaxFormattedTitle = sTitle.Replace("{0}", "APW");
            //}

        }
        catch (Exception e)
        {
           Assert.Fail(e.Message);
        }
        return sSotaxFormattedTitle;
    }
        public static void VerifyPaneExist(Window sWindow, String sParentPanel, String sChildPanel = null, String sSubChild = null)
        {
           
            bool ChildExist, ChildExistID, ChildExistName;
            try
            {

                Panel Pane = sWindow.Get<Panel>(SearchCriteria.ByAutomationId(sParentPanel));
                if (sChildPanel != null)
                {
                    ChildExistID = Pane.Exists(SearchCriteria.ByAutomationId(sChildPanel));
                    
                    if (sSubChild != null)
                    {
                        ChildExistName = Pane.Exists(SearchCriteria.ByText(sChildPanel));
                        if ((ChildExistName == true) || (ChildExistID == true))
                            ChildExist = true;
                        else
                            ChildExist = false;

                        Assert.IsTrue(ChildExist.Equals(true), "<B> Expected location for " + sChildPanel + " is </B> " + sParentPanel);
                    }
                }
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

                public static Boolean VerifyFieldCheck(Window sWindow)
        {
            
            try
            {
                
                SearchCriteria searchCriteria = SearchCriteria.ByAutomationId("").AndControlType(ControlType.Pane);
                Panel Obj= sWindow.Get<Panel>(""); //default search mechanism is by automation id
                Obj = sWindow.Get<Panel>(searchCriteria);

                if (Obj != null)
                    return true;
                else
                    return false;


            }
            catch (Exception ex)
            {
                Console.WriteLine("No errors found, entered value is correct", ex.ToString());
                return false;
            }
        }
        
        internal static void ClickExport(string sExportPath)
        {
            throw new NotImplementedException();
        }
    }
}


