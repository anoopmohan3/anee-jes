using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using TestStack.White.UIItems.Finders;
using TestStack.White.UIItems.WindowItems;
//using White.Support_Files;
using TestStack.White.UIItems;
using TestStack.White.UIItems.ListBoxItems;
using TestStack.White.Configuration;
using System.Text.RegularExpressions;
using System.IO;
using System.Text;
using System.Diagnostics;
using TestStack.White;
using TestStack.White.Utility;
using System.Linq;
using TestStack.White.UIItems.TabItems;
using TestStack.White.InputDevices;
using TestStack.White.WindowsAPI;
using TestStack.White.UIItems.TableItems;
using TestStack.White.UIItems.TreeItems;
using System.Windows.Automation;
using TestStack.White.UIItems.WPFUIItems;

namespace Whiteclass
{
    class CommonMethods
    {

        public static string DesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public static string sExpImpPath = CommonMethods.DesktopPath + @"\ExportImportMethods\";
        public static string TareFilepath = CommonMethods.DesktopPath + @"\ExportImportMethods\";
        public static string SaveTarepath = CommonMethods.DesktopPath + @"\ExportImportMethods\";
        /** Function to select Verify text*/

        public static void selectTabItem(Window sWindow, string sTabControl, String sTabItem)
        {
            try
            {
                Tab sTabCtrl = sWindow.Get<Tab>(SearchCriteria.ByAutomationId(sTabControl));

                sTabCtrl.SelectTabPage(sTabItem);
                //Common.Sleep(Common.obj_st_GlobalTestSupport.iSleepTime);
                
                Assert.IsTrue(sTabCtrl.SelectedTab.ToString().Equals(sTabItem), "Verify Selected Tab item is " + sTabItem);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        public static void ClickButton(Window sWindow, String sButtonID, int iDepth = 0, String sLogInfo = null)
        {
            try
            {
                CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = iDepth;
                Button sButton;
                char caseSwitch = (char)Common.ObjectIdentifier[sButtonID];
                switch (caseSwitch)
                {
                    case '#':
                        sButton = (Button)sWindow.Get(SearchCriteria.ByText(sButtonID), TimeSpan.FromSeconds(15));
                        break;
                    case '?':
                        sButton = (Button)sWindow.Get(SearchCriteria.ByClassName(sButtonID), TimeSpan.FromSeconds(15));
                        break;
                    case '$':
                        sButton = (Button)sWindow.Get(SearchCriteria.ByAutomationId(sButtonID), TimeSpan.FromSeconds(15));
                        break;
                    default:
                        sButton = (Button)sWindow.Get(SearchCriteria.ByAutomationId(sButtonID), TimeSpan.FromSeconds(15));
                        break;
                }

                
                
                    if (sButton.Name != "" && sButton.Name != null)
                    {
                        //ATEL.LogTestSteps("Click on '" + sButton.Name + "' Button", 2);
                    }
                    else if (sButton.Text != "" && sButton.Text != null)
                    {
                        //ATEL.LogTestSteps("Click on '" + sButton.Text + "' Button", 2);
                    }
                

                Assert.IsTrue(sButton.Visible, "Button is not currently Visible");
                sButton.Click();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

            }
            catch (Exception e)
            {
                
                Assert.Fail(e.Message);
            }
        }

        internal static void doubleClickTableRow(Window sOTAXAddReagentWindow, string v)
        {
            throw new NotImplementedException();
        }

        internal static void selectTabItem(Window methodSetupWindow, string v)
        {
            throw new NotImplementedException();
        }



        /// <summary>
        /// ClickCheckBox 
        /// </summary>
        /// <param name="sWindow"></param>
        /// <param name="sCheckboxID"></param>
        /// <param name="sDepth"></param>
        /// <param name="sLogInfo"></param>
        public static void radioButtonEnabled(Window sWindow, String sCheckboxID, int sDepth = 0, String sLogInfo = null)
        {
            try
            {
                CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = sDepth;
                CheckBox sCheckbox;
                char caseSwitch = (char)Common.ObjectIdentifier[sCheckboxID];
                switch (caseSwitch)
                {
                    case '#':
                        sCheckbox = (CheckBox)sWindow.Get(SearchCriteria.ByText(sCheckboxID), TimeSpan.FromSeconds(15));
                        break;
                    case '?':
                        sCheckbox = (CheckBox)sWindow.Get(SearchCriteria.ByClassName(sCheckboxID), TimeSpan.FromSeconds(15));
                        break;
                    case '$':
                        sCheckbox = (CheckBox)sWindow.Get(SearchCriteria.ByAutomationId(sCheckboxID), TimeSpan.FromSeconds(15));
                        break;
                    default:
                        sCheckbox = (CheckBox)sWindow.Get(SearchCriteria.ByAutomationId(sCheckboxID), TimeSpan.FromSeconds(15));
                        break;
                }

                
                
                    if (sCheckbox.Name != "" && sCheckbox.Name != null)
                    {
                        //ATEL.LogTestSteps("Click on '" + sCheckbox.Name + "' Button", 2);
                    }
                    else if (sCheckbox.Text != "" && sCheckbox.Text != null)
                    {
                        //ATEL.LogTestSteps("Click on '" + sCheckbox.Text + "' Button", 2);
                    }
                

                Assert.IsTrue(sCheckbox.Visible, "CheckBox is not currently Visible");
                sCheckbox.Click();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        internal static void chkBoxSelect(Window sotaxUserInformationWindow, string v)
        {
            throw new NotImplementedException();
        }

        public static Boolean verifyButtonText(Window sWindow, String sButtonID, String sExpectedBtnText)
        {
            Boolean bFound = false;
            try
            {
                Button sButton;

                char caseSwitch = (char)Common.ObjectIdentifier[sButtonID];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sButton = sWindow.Get<Button>(SearchCriteria.ByText(sButtonID));
                        break;
                    case '?':
                        sButton = sWindow.Get<Button>(SearchCriteria.ByClassName(sButtonID));
                        break;
                    case '$':
                        sButton = sWindow.Get<Button>(SearchCriteria.ByAutomationId(sButtonID));
                        break;
                    default:
                        sButton = sWindow.Get<Button>(SearchCriteria.ByAutomationId(sButtonID));
                        break;
                }
                //ATEL.LogTestSteps("Verify Button '" + sButtonID + "' text is set to '" + sExpectedBtnText + "'", 1);
                // Assert.IsTrue(sButton.Text.Equals(sExpectedBtnText) || sButton.Name.Equals(sExpectedBtnText), "Expected Value : " + sExpectedBtnText + " Current Text Value: " + sButton.Text + "Current Name Value" + sButton.Name);

                if (!sButton.Name.Equals(""))
                {
                    Assert.IsTrue(sButton.Name.Equals(sExpectedBtnText), "Expected Value : " + sExpectedBtnText + " Current Value: " + sButton.Name);
                }
                else if (!sButton.Text.Equals(""))
                {
                    Assert.IsTrue(sButton.Text.Equals(sExpectedBtnText), "Expected Value : " + sExpectedBtnText + " Current Value: " + sButton.Text);
                }
                else
                {
                    Assert.Fail("Both Name & Text is not set for " + sExpectedBtnText);
                }
                if (sButton.Name.Equals(sExpectedBtnText))
                    bFound = true;
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return bFound;
        }
        
        
        public static void ButtonSetFocus(Window sWindow, String sButtonID)
        {
            try
            {

                Button sButton;

                char caseSwitch = (char)Common.ObjectIdentifier[sButtonID];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sButton = sWindow.Get<Button>(SearchCriteria.ByText(sButtonID));
                        break;
                    case '?':
                        sButton = sWindow.Get<Button>(SearchCriteria.ByClassName(sButtonID));
                        break;
                    case '$':
                        sButton = sWindow.Get<Button>(SearchCriteria.ByAutomationId(sButtonID));
                        break;
                    default:
                        sButton = sWindow.Get<Button>(SearchCriteria.ByAutomationId(sButtonID));
                        break;
                }

                //ATEL.LogTestSteps("Set Focus on 'Button" + sButton.Text);
                sButton.Focus();

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Verify Button is Disabled
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="searchCriteria">Button text to check it's Disabled</param>
        public static void verifyButtonIsNotVisible(Window sWindow, String searchCriteria, string sbuttonText = null)
        {
            try
            {
                char caseSwitch = Common.identifier;// Convert.ToChar(sField.Substring(0, 1));
                Button button;
                switch (caseSwitch)
                {
                    case '#':
                        button = sWindow.Get<Button>(SearchCriteria.ByText(searchCriteria));
                        break;
                    case '?':
                        button = sWindow.Get<Button>(SearchCriteria.ByClassName(searchCriteria));
                        break;
                    case '$':
                        button = sWindow.Get<Button>(SearchCriteria.ByAutomationId(searchCriteria));
                        break;
                    default:
                        button = sWindow.Get<Button>(SearchCriteria.ByAutomationId(searchCriteria));
                        break;
                }

                //if (sbuttonText == null)
                    //ATEL.LogTestSteps("Verify Button" + button.Text + " is not visible", 2);
                //else
                    //ATEL.LogTestSteps("Verify Button" + sbuttonText + " is not visible", 2);

                Assert.IsFalse(button.Visible, "Button is currently Visible");
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Verify Button is Enabled
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="searchCriteria">Button text to check it's Enabled</param>
        public static void verifyButtonEnabled(Window sWindow, String searchCriteria, int sDepth = 0)
        {
            try
            {

                char caseSwitch = Common.identifier;// Convert.ToChar(sField.Substring(0, 1));
                Button button;
                //ATEL.LogTestSteps("Verify Button : " + searchCriteria + " is Enabled", 2);

                switch (caseSwitch)
                {
                    case '#':
                        button = (Button)sWindow.Get(SearchCriteria.ByText(searchCriteria));
                        break;
                    case '?':
                        button = (Button)sWindow.Get(SearchCriteria.ByClassName(searchCriteria));
                        break;
                    case '$':
                        button = (Button)sWindow.Get(SearchCriteria.ByAutomationId(searchCriteria));
                        break;
                    default:
                        button = (Button)sWindow.Get(SearchCriteria.ByAutomationId(searchCriteria));
                        break;
                }

                sWindow.WaitWhileBusy();
                Assert.IsTrue(button.Enabled, "Button is not 'Enabled'");

                CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        internal static void getList(Window sOTAXReagentWindow, string v)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Verify Button is Disabled
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="searchCriteria">Button text to check it's Disabled</param>
        public static void verifyButtonDisabled(Window sWindow, String searchCriteria, int sDepth = 0)
        {
            try
            {
                CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = sDepth;

                char caseSwitch = Common.identifier;// Convert.ToChar(sField.Substring(0, 1));
                Button button;
                //ATEL.LogTestSteps("Verify Button : " + searchCriteria + " is Disabled", 2);

                switch (caseSwitch)
                {
                    case '#':
                        button = sWindow.Get<Button>(SearchCriteria.ByText(searchCriteria));
                        break;
                    case '?':
                        button = sWindow.Get<Button>(SearchCriteria.ByClassName(searchCriteria));
                        break;
                    case '$':
                        button = sWindow.Get<Button>(SearchCriteria.ByAutomationId(searchCriteria));
                        break;
                    default:
                        button = sWindow.Get<Button>(SearchCriteria.ByAutomationId(searchCriteria));
                        break;
                }
                Assert.IsFalse(button.Enabled, "Verify Button is Disabled");
                CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Function to verify Column Headers contain expected text in header also return column headers
        /// </summary>
        /// <param name="sWindow">Window Declaration</param>
        /// <param name="sDataGridAutoID">DataGrid Declaration</param>
        /// <param name="sHeaderName">Name of the Data Table List</param>
        /// <param name="sHeaderList"> List of Headers to verify in DataGrid/Table</param>
        /// <returns>Returns list of column headers</returns>
        public static List<string> VerifyColumnHeadersContain(Window sWindow, String sDataGrid, string sHeaderName, List<String> expectedHeaderList = null)
        {
            List<string> currentHeaderList = new List<string>();
            try
            {
                ListView getDataGridList;
                char caseSwitch = (char)Common.ObjectIdentifier[sDataGrid];
                switch (caseSwitch)
                {
                    case '#':
                        getDataGridList = sWindow.Get<ListView>(SearchCriteria.ByText(sDataGrid));
                        break;
                    case '?':
                        getDataGridList = sWindow.Get<ListView>(SearchCriteria.ByClassName(sDataGrid));
                        break;
                    case '$':
                        getDataGridList = sWindow.Get<ListView>(SearchCriteria.ByAutomationId(sDataGrid));
                        break;
                    default:
                        getDataGridList = sWindow.Get<ListView>(SearchCriteria.ByAutomationId(sDataGrid));
                        break;
                }

                //ATEL.LogTestSteps("Verify Column Header '" + sHeaderName + "' List");

                //Get List of Headers
                for (int i = 0; i < getDataGridList.Header.Columns.Count; i++)
                {
                    currentHeaderList.Add(Regex.Replace(getDataGridList.Header.Columns[i].Text.ToString(), @"\r\n|\r|\n", " "));
                }

                //Compare Expected Header List Columns with Current List of Comumns
                if (expectedHeaderList != null)
                {
                    for (int i = 0; i < expectedHeaderList.Count; i++)
                    {
                        //ATEL.LogTestSteps("Verify " + (i + 1) + " Column header for '" + sHeaderName + "' contains '" + expectedHeaderList[i].ToString() + "'", 1);
                        Assert.IsTrue((currentHeaderList[i].ToString().Contains(expectedHeaderList[i].ToString())), "<B>Expected Header Text</B>" + expectedHeaderList[i].ToString() + " <B>Current Header Text:</B>" + currentHeaderList[i].ToString());
                    }
                }
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return currentHeaderList;
        }


        public static void verifyLabelEnabled(Window sWindow, String lblID)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[lblID];// Convert.ToChar(sField.Substring(0, 1));
                Label lblTextID;

                switch (caseSwitch)
                {
                    case '#':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByText(lblID));
                        break;
                    case '?':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByClassName(lblID));
                        break;
                    case '$':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                    default:
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                }
                Assert.IsTrue(lblTextID.Enabled, " Label " + lblTextID.Text + " is currently Disabled");
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        public static void verifyLabelDisabled(Window sWindow, String lblID)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[lblID];// Convert.ToChar(sField.Substring(0, 1));
                Label lblTextID;

                switch (caseSwitch)
                {
                    case '#':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByText(lblID));
                        break;
                    case '?':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByClassName(lblID));
                        break;
                    case '$':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                    default:
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                }
                Assert.IsFalse(lblTextID.Enabled, " Label " + lblTextID.Text + " is currently enabled");
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Verify Test 
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="sTextBoxID">TextBox Automation ID</param>
        /// <param name="sExpectedText">Expecte value for Text Box</param>
        /// <param name="iIndex">Index Value: Provide Location only when there is no unique Identified.</param>
        public static void verifyText(Window sWindow, String sTextBoxID, String sExpectedText, int iIndex = 0)
        {

            try
            {
                TextBox sTextBox;

                char caseSwitch = (char)Common.ObjectIdentifier[sTextBoxID];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByText(sTextBoxID).AndIndex(iIndex));
                        break;
                    case '?':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByClassName(sTextBoxID).AndIndex(iIndex));
                        break;
                    case '$':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(sTextBoxID).AndIndex(iIndex));
                        break;
                    default:
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(sTextBoxID).AndIndex(iIndex));
                        break;
                }
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                //ATEL.LogTestSteps("Verify value for " + sTextBox.Name + " is set to '" + sExpectedText + "'", 1);
                Assert.IsTrue(sTextBox.Text.Equals(sExpectedText), "Expected Value : " + sExpectedText + " Current Value: " + sTextBox.Text);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }


        /// <summary>
        /// User this verifyText when Label name not assigned. 
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="sTextBoxID">TextBox Automation ID</param>
        /// <param name="sLabel">Label for the Text Box</param>
        /// <param name="sExpectedText">Expecte value for Text Box</param>
        /// <param name="iIndex">Index Value: Provide Location only when there is no unique Identified.</param>
        public static void verifyText(Window sWindow, String sTextBoxID, String sLabel, String sExpectedText, int iIndex = 0)
        {
            try
            {
                TextBox sTextBox;

                char caseSwitch = (char)Common.ObjectIdentifier[sTextBoxID];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByText(sTextBoxID).AndIndex(iIndex));
                        break;
                    case '?':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByClassName(sTextBoxID).AndIndex(iIndex));
                        break;
                    case '$':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(sTextBoxID).AndIndex(iIndex));
                        break;
                    default:
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(sTextBoxID).AndIndex(iIndex));
                        break;
                }
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                //ATEL.LogTestSteps("Verify value for " + sLabel + " is set to '" + sExpectedText + "'", 1);
                Assert.IsTrue(sTextBox.Text.Equals(sExpectedText), "Expected Value : " + sExpectedText + " Current Value: " + sTextBox.Text);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Verify Text is Enabled
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="sTextBoxID">TextBox Automation ID</param>
        /// <param name="iIndex">Index Value: Provide Location only when there is no unique Identified.</param>
        public static void verifyTextEnabled(Window sWindow, String sTextBoxID, int iIndex = 0)
        {

            try
            {
                TextBox sTextBox;

                char caseSwitch = (char)Common.ObjectIdentifier[sTextBoxID];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByText(sTextBoxID).AndIndex(iIndex));
                        break;
                    case '?':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByClassName(sTextBoxID).AndIndex(iIndex));
                        break;
                    case '$':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(sTextBoxID).AndIndex(iIndex));
                        break;
                    default:
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(sTextBoxID).AndIndex(iIndex));
                        break;
                }
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                //ATEL.LogTestSteps("Verify text " + sTextBox.Name + " is Enabled'", 1);
                Assert.IsTrue(sTextBox.Enabled, "Failed at : Verify " + sTextBox.Text + " is enabled");

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Verify Expected Text within a Group Box
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="sTextBoxID">TextBox Automation ID</param>
        /// <param name="sExpectedText">Expected Text</param>
        public static void verifyTextGroupBox(Window sWindow, String sTextBoxID, String sExpectedText)
        {
            try
            {
                GroupBox sGroupBoxText;

                char caseSwitch = (char)Common.ObjectIdentifier[sTextBoxID];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sGroupBoxText = sWindow.Get<GroupBox>(SearchCriteria.ByText(sTextBoxID));
                        break;
                    case '?':
                        sGroupBoxText = sWindow.Get<GroupBox>(SearchCriteria.ByClassName(sTextBoxID));
                        break;
                    case '$':
                        sGroupBoxText = sWindow.Get<GroupBox>(SearchCriteria.ByAutomationId(sTextBoxID));
                        break;
                    default:
                        sGroupBoxText = sWindow.Get<GroupBox>(SearchCriteria.ByAutomationId(sTextBoxID));
                        break;
                }
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                //ATEL.LogTestSteps("Verify GroupBox " + sTextBoxID + " text is set to '" + sExpectedText + "'");

                Assert.IsTrue(sGroupBoxText.Name.Equals(sExpectedText), "Expected Value : " + sExpectedText + " Current Value: " + sGroupBoxText.Name);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        ///  Set Text 
        /// </summary>
        /// <param name="sWindow">Window ID</param>
        /// <param name="stxtboxfield">Text Box ID</param>
        /// <param name="sTextBoxName">Text Box Name</param>
        /// <param name="sSetText">set text Value</param>
        public static void setText(Window sWindow, String stxtboxfield, String sTextBoxName, String sSetText)
        {
            try
            {
                TextBox sTextBox;
                //ATEL.LogTestSteps("Set text for  ' " + sTextBoxName + "' to " + sSetText, 2);

                char caseSwitch = (char)Common.ObjectIdentifier[stxtboxfield];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByText(stxtboxfield));
                        break;
                    case '?':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByClassName(stxtboxfield));
                        break;
                    case '$':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(stxtboxfield));
                        break;
                    default:
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(stxtboxfield));
                        break;
                }
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

                Assert.IsTrue(sTextBox.Visible, sTextBoxName + " is not visible");
                sTextBox.SetValue(sSetText);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        public static void selectComboBoxListItem(Window sWindow, String sComboBox, String sOption)
        {

            try
            {
                ComboBox comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                comboBox.Focus();
                comboBox.Click();
                Keyboard.Instance.Enter(sOption);
                sWindow.ReloadIfCached();

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }


        /// <summary>
        /// Verify Text is Disabled
        /// </summary>
        /// <param name="sWindow">Window ID</param>
        /// <param name="stxtboxfield">Text Box ID</param>
        /// <param name="sTextBoxName">Text Box Name</param>
        /// <param name="sSetText">set text Value</param>
        public static void verifyTextDisabled(Window sWindow, String stxtboxfield, string sLog = "")
        {
            try
            {
                TextBox sTextBox;

                char caseSwitch = (char)Common.ObjectIdentifier[stxtboxfield];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByText(stxtboxfield));
                        break;
                    case '?':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByClassName(stxtboxfield));
                        break;
                    case '$':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(stxtboxfield));
                        break;
                    default:
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(stxtboxfield));
                        break;
                }
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                
                    //ATEL.LogTestSteps("Verify Text is  disabled", 2);

                Assert.IsFalse(sTextBox.Enabled, "Verify Text is disabled");

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        /// <summary>
        /// Set Text is Enabled
        /// </summary>
        /// <param name="sWindow">Window ID</param>
        /// <param name="stxtboxfield">Text Box ID</param>
        /// <param name="sTextBoxName">Text Box Name</param>
        /// <param name="sSetText">set text Value</param>
        public static void verifyTexEnabled(Window sWindow, String stxtboxfield, string sLog = "")
        {
            try
            {
                TextBox sTextBox;

                char caseSwitch = (char)Common.ObjectIdentifier[stxtboxfield];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByText(stxtboxfield));
                        break;
                    case '?':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByClassName(stxtboxfield));
                        break;
                    case '$':
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(stxtboxfield));
                        break;
                    default:
                        sTextBox = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(stxtboxfield));
                        break;
                }
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                
                    //ATEL.LogTestSteps("Verify Text is  Enabled", 2);

                Assert.IsTrue(sTextBox.Enabled, "Verify Text is Enabled");

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }


        /// <summary>
        /// Verify Text
        /// </summary>
        /// <param name="sWindow"></param>
        /// <param name="sField"></param>
        public static void VerifyText(Window sWindow, String sField)
        {
            try
            {
                char caseSwitch = Common.identifier;// Convert.ToChar(sField.Substring(0, 1));
                Label txtResult;

                //ATEL.LogTestSteps("Verify Label:" + sField + " Exists", 2);
                switch (caseSwitch)
                {
                    case '#':
                        txtResult = sWindow.Get<Label>(SearchCriteria.ByText(sField));
                        break;
                    case '?':
                        txtResult = sWindow.Get<Label>(SearchCriteria.ByClassName(sField));
                        break;
                    case '$':
                        txtResult = sWindow.Get<Label>(SearchCriteria.ByAutomationId(sField));
                        break;
                    default:
                        txtResult = sWindow.Get<Label>(SearchCriteria.ByAutomationId(sField));
                        break;
                }

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Verify Label is Visible
        /// </summary>
        /// <param name="sWindow"></param>
        /// <param name="lblID"></param>
        /// <param name="sExpectedText"></param>
        public static void verifyLabelVisible(Window sWindow, String lblID, String sExpectedText)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[lblID];// Convert.ToChar(sField.Substring(0, 1));
                Label lblTextID;

                switch (caseSwitch)
                {
                    case '#':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByText(lblID));
                        break;
                    case '?':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByClassName(lblID));
                        break;
                    case '$':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                    default:
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                }
                lblTextID.Focus();

                //ATEL.LogTestSteps("Verify Label '" + sExpectedText + " ' is Visible.", 1);
                Assert.IsTrue(lblTextID.Visible, "Label '" + sExpectedText + "' is not visible");
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Verify Label is Visible
        /// </summary>
        /// <param name="sWindow"></param>
        /// <param name="lblID"></param>
        /// <param name="sExpectedText"></param>
        public static void verifyLabelNotVisible(Window sWindow, String lblID, String sExpectedText)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[lblID];// Convert.ToChar(sField.Substring(0, 1));
                Label lblTextID;

                switch (caseSwitch)
                {
                    case '#':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByText(lblID));
                        break;
                    case '?':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByClassName(lblID));
                        break;
                    case '$':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                    default:
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                }
                lblTextID.Focus();

                //ATEL.LogTestSteps("Verify Label '" + sExpectedText + " ' is not Visible.", 1);
                Assert.IsFalse(lblTextID.Visible, "Label '" + sExpectedText + "' is visible it should not be visible");
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Use this Function to Verify Text for the lable
        /// </summary>
        /// <param name="sWindow">Provide Window declaration</param>
        /// <param name="slblObject">Provide label ID to get the text. If it's automationID then it will go to any sub level</param>
        /// <param name="sExpectedText">Provide Expected text</param>
        public static void verifyLabel(Window sWindow, String lblID, String sExpectedText)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[lblID];// Convert.ToChar(sField.Substring(0, 1));
                Label lblTextID;

                switch (caseSwitch)
                {
                    case '#':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByText(lblID));
                        break;
                    case '?':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByClassName(lblID));
                        break;
                    case '$':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                    default:
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                }
                lblTextID.Focus();

                sExpectedText = (Regex.Replace(sExpectedText.Trim(), @"\s+", " "));
                string sCurrentDescription = "";

                //ATEL.LogTestSteps("Verify Label '" + sExpectedText + "' Exists.", 1);

                if (!lblTextID.Name.Equals(""))
                {
                    sCurrentDescription = (Regex.Replace(lblTextID.Name.Trim(), @"\s+", " "));
                    Assert.IsTrue(sCurrentDescription.Equals(sExpectedText), "Expected Value : " + sExpectedText + " Current Value: " + sCurrentDescription);
                }
                else if (!lblTextID.Text.Equals(""))
                {
                    sCurrentDescription = (Regex.Replace(lblTextID.Text.Trim(), @"\s+", " "));
                    Assert.IsTrue(sCurrentDescription.Equals(sExpectedText), "Expected Value : " + sExpectedText + " Current Value: " + sCurrentDescription);
                }
                else
                {
                    Assert.Fail("Both Name & Text is not set for " + lblTextID);
                }

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Use this Function to Verify Text contains text provided for the lable
        /// </summary>
        /// <param name="sWindow">Provide Window declaration</param>
        /// <param name="slblObject">Provide label ID to get the text. If it's automationID then it will go to any sub level</param>
        /// <param name="sExpectedText">Provide Expected text</param>
        public static Boolean verifyLabelContainsText(Window sWindow, String lblID, String sExpectedText)
        {
            Boolean bFound = false;

            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[lblID];// Convert.ToChar(sField.Substring(0, 1));
                Label lblTextID;

                switch (caseSwitch)
                {
                    case '#':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByText(lblID));
                        break;
                    case '?':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByClassName(lblID));
                        break;
                    case '$':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                    default:
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                }
                lblTextID.Focus();

                sExpectedText = (Regex.Replace(sExpectedText.Trim(), @"\s+", " "));
                string sCurrentDescription = "";


                //ATEL.LogTestSteps("Verify Label '" + sExpectedText + "' Exists.", 1);

                if (!lblTextID.Name.Equals(""))
                {
                    sCurrentDescription = (Regex.Replace(lblTextID.Name.Trim(), @"\s+", " "));
                    Assert.IsTrue(sCurrentDescription.Contains(sExpectedText), "Substring of Expected Value: " + sExpectedText + " Current Value: " + sCurrentDescription);
                }
                else if (!lblTextID.Text.Equals(""))
                {
                    sCurrentDescription = (Regex.Replace(lblTextID.Text.Trim(), @"\s+", " "));
                    Assert.IsTrue(sCurrentDescription.Contains(sExpectedText), "Substring of Expected Value : " + sExpectedText + " Current Value: " + sCurrentDescription);
                }
                else
                {
                    Assert.Fail("Both Name & Text is not set for " + lblTextID);
                }

                if (sCurrentDescription.Contains(sExpectedText))
                    bFound = true;

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return bFound;
        }

        /// <summary>
        /// Use this Function to Verify Text for the lable
        /// </summary>
        /// <param name="sWindow">Provide Window declaration</param>
        /// <param name="slblObject">Provide label ID to get the text. If it's automationID then it will go to any sub level</param>
        /// <param name="sExpectedText">Provide Expected text</param>
        public static string getLabelText(Window sWindow, String lblID)
        {
            string sLabelText = "";

            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[lblID];// Convert.ToChar(sField.Substring(0, 1));
                Label lblTextID;
                switch (caseSwitch)
                {
                    case '#':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByText(lblID));
                        break;
                    case '?':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByClassName(lblID));
                        break;
                    case '$':
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                    default:
                        lblTextID = sWindow.Get<Label>(SearchCriteria.ByAutomationId(lblID));
                        break;
                }
                lblTextID.Focus();
                if (!lblTextID.Name.Equals(""))
                {
                    sLabelText = lblTextID.Name.Trim();
                }
                else if (!lblTextID.Text.Equals(""))
                {
                    sLabelText = lblTextID.Text.Trim();
                }
                else
                {
                    sLabelText = "";
                    Assert.Fail("Both Name & Text is not set for " + lblTextID);
                }
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return sLabelText;
        }

        public static string getTextValue(Window sWindow, String txtID)
        {
            string sLabelText = "";

            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[txtID];// Convert.ToChar(sField.Substring(0, 1));
                TextBox txtTextID;
                switch (caseSwitch)
                {
                    case '#':
                        txtTextID = sWindow.Get<TextBox>(SearchCriteria.ByText(txtID));
                        break;
                    case '?':
                        txtTextID = sWindow.Get<TextBox>(SearchCriteria.ByClassName(txtID));
                        break;
                    case '$':
                        txtTextID = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(txtID));
                        break;
                    default:
                        txtTextID = sWindow.Get<TextBox>(SearchCriteria.ByAutomationId(txtID));
                        break;
                }
                txtTextID.Focus();
                sLabelText = txtTextID.Text.ToString();
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return sLabelText;
        }
        public static GroupBox getGroupBox(Window sWindow, String sGroupBox)
        {
            //Default Value Set to Automaton ID
            GroupBox groupBox = null;
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[sGroupBox]; // Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        groupBox = sWindow.Get<GroupBox>(SearchCriteria.ByText(sGroupBox));
                        break;
                    case '?':
                        groupBox = sWindow.Get<GroupBox>(SearchCriteria.ByClassName(sGroupBox));
                        break;
                    case '$':
                        groupBox = sWindow.Get<GroupBox>(SearchCriteria.ByAutomationId(sGroupBox));
                        break;
                    default:
                        groupBox = sWindow.Get<GroupBox>(SearchCriteria.ByAutomationId(sGroupBox));
                        break;
                }
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return groupBox;
        }

        public static ComboBox getComboBox(Window sWindow, String sComboBox)
        {
            //Default Value Set to Automaton ID
            ComboBox comboBox = null;
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[sComboBox]; // Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByText(sComboBox));
                        break;
                    case '?':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByClassName(sComboBox));
                        break;
                    case '$':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                    default:
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                }
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return comboBox;
        }

        public static ListView getList(GroupBox sObject, String sList)
        {
            //Default Value Set to Automaton ID
            ListView sListView = null;
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[sList]; // Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sListView = sObject.Get<ListView>(SearchCriteria.ByText(sList));
                        break;
                    case '?':
                        sListView = sObject.Get<ListView>(SearchCriteria.ByClassName(sList));
                        break;
                    case '$':
                        sListView = sObject.Get<ListView>(SearchCriteria.ByAutomationId(sList));
                        break;
                    default:
                        sListView = sObject.Get<ListView>(SearchCriteria.ByAutomationId(sList));
                        break;
                }

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return sListView;
        }

        public static void verifyComboList(Window sWindow, string sComboBox, List<string> sExpectedList)
        {
            try
            {
                ComboBox comboBox;
                char caseSwitch = (char)Common.ObjectIdentifier[sComboBox]; // Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)

                {
                    case '#':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByText(sComboBox));
                        break;
                    case '?':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByClassName(sComboBox));
                        break;
                    case '$':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                    default:
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                }

                // List<string> verifyList = verifyList = new List<String> { pg_EXTRACTSettingsPage.LanguageOption01(), pg_EXTRACTSettingsPage.LanguageOption02() };


                for (int i = 0; i < comboBox.Items.Count; i++)
                {
                    //ATEL.LogTestSteps("Verify '" + i + "' list Item is set to " + sExpectedList[i].ToString(), 1);
                    Assert.IsTrue((comboBox.Items[i].Text).Equals(sExpectedList[i].ToString()), "Expected: " + i + " value : " + sExpectedList[i].ToString() + " Current Value: " + comboBox.Items[i].Text);
                }
                sWindow.Focus();
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        internal static void ButtonSetFocus(Window saveTareFileWindow, string v1, string v2, string v3)
        {
            throw new NotImplementedException();
        }

        public static void selectComboList(Window sWindow, string sComboBox, string sSelectListItem, string sLogInfo = null)
        {
            try
            {
                ComboBox comboBox;
                char caseSwitch = (char)Common.ObjectIdentifier[sComboBox];
                switch (caseSwitch)

                {
                    case '#':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByText(sComboBox));
                        break;
                    case '?':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByClassName(sComboBox));
                        break;
                    case '$':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                    default:
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                }

                sWindow.Focus();

                if (comboBox.Name != "" && comboBox.Name != null)
                    {
                        //ATEL.LogTestSteps("Select '" + sSelectListItem + " from combo box" + comboBox.Name, 2);
                    }
                

                comboBox.Select(sSelectListItem);
                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }
        /// <summary>
        /// Select List Item from the datagrid  
        /// </summary>
        /// <param name="sWindow">Window</param>
        /// <param name="sDataGridID">DataGrid ID</param>
        /// <param name="iRowNumber">Row number of the item to be selected</param>
        public static void selectListItem(Window sWindow, string sDataGridID, int iRowNumber)
        {
            try
            {
                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                
                ListBox list = sWindow.Get<ListBox>(SearchCriteria.ByAutomationId(sDataGridID));
                list.Focus();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                list.Items[iRowNumber].Focus();
                list.Items[iRowNumber].Select();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

            }

            catch (Exception e)
            {
               Assert.Fail(e.Message);
            }

        }

        public static void verifyListItemText(Window sWindow, String sPane, String sActivityText)
        {
            try
            {
                Boolean bfound = false;
                var listView = sWindow.Get<ListView>(SearchCriteria.ByAutomationId(sPane));
                string sCurrentDescription = "";
                for (int i = 0; i < listView.Items.Count; i++)
                {
                    if (listView.Rows[i].Cells[1].ToString().Contains(sActivityText))
                    {
                        bfound = true;
                        sCurrentDescription = listView.Rows[i].Cells[2].Name.ToString();
                        sCurrentDescription = (Regex.Replace(sCurrentDescription.Trim(), @"\s+", " "));
                        //sExpectedDescriptionText = (Regex.Replace(sExpectedDescriptionText.Trim(), @"\s+", " "));

                        //Assert.IsTrue(sCurrentDescription.Equals(sExpectedDescriptionText), "<B>Expected Description</B>" + sExpectedDescriptionText + "<B> Current Description </B>" + sCurrentDescription);
                       // if (sSchedule != "")
                       // {
                            //string sCurrentSchedule = listView.Rows[i].Get<Label>(SearchCriteria.ByAutomationId(pg_COREMaintenancePage.lblScheduleID())).Name;
                        //    Assert.IsTrue(sCurrentSchedule.Equals(sSchedule), "<B>Expected Schedule</B>" + sSchedule + "<B> Current Schedule </B>" + sCurrentSchedule);
                        //}
                        break;
                    }
                }
                if (bfound == false)
                {
                    Assert.Fail(sActivityText + " Not found");
                }
            }
            catch (Exception e)
            {
               Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Select List Item 
        /// </summary>
        /// <param name="sWindow">Window</param>
        /// <param name="sListID">List ID</param>
        /// <param name="sItemText">Item Text</param>
        /// <param name="sDepth">Depth</param>
        /// <param name="sLogInfo">Add Log</param>
        public static void SelectListItem(Window sWindow, String sListID, String sItemText, int sDepth = 0, String sLogInfo = null)
        {
            try
            {
                CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = sDepth;
                ListBox sListBox;
                char caseSwitch = (char)Common.ObjectIdentifier[sListID];
                switch (caseSwitch)
                {
                    case '#':
                        sListBox = (ListBox)sWindow.Get(SearchCriteria.ByText(sListID), TimeSpan.FromSeconds(15));
                        break;
                    case '?':
                        sListBox = (ListBox)sWindow.Get(SearchCriteria.ByClassName(sListID), TimeSpan.FromSeconds(15));
                        break;
                    case '$':
                        sListBox = (ListBox)sWindow.Get(SearchCriteria.ByAutomationId(sListID), TimeSpan.FromSeconds(15));
                        break;
                    default:
                        sListBox = (ListBox)sWindow.Get(SearchCriteria.ByAutomationId(sListID), TimeSpan.FromSeconds(15));
                        break;
                }

                if (sListBox.Name != "" && sListBox.Name != null)
                    {
                        //ATEL.LogTestSteps("Select on '" + sListBox.Name + "' Item", 2);
                    }
                
                Assert.IsTrue(sListBox.Visible, "List Box is not currently Visible");
                //ATEL.LogTestSteps("Select list item " + sItemText, 1);
                sListBox.Items.Select(sItemText);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }



        /// <summary>
        /// Verify Combo Box is Disabled
        /// </summary>
        /// <param name="sWindow">Window</param>
        /// <param name="sComboBox">Combo Box ID</param>
        /// <param name="sLogInfo">Log test Setp</param>
        public static void verifyComboListDisabled(Window sWindow, string sComboBox, string sLogInfo = null)
        {
            try
            {
                ComboBox comboBox;
                char caseSwitch = (char)Common.ObjectIdentifier[sComboBox];
                switch (caseSwitch)

                {
                    case '#':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByText(sComboBox));
                        break;
                    case '?':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByClassName(sComboBox));
                        break;
                    case '$':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                    default:
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                }

                sWindow.Focus();

                //ATEL.LogTestSteps("Verify 'Combo Box' is disabled", 2);

                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Assert.IsFalse(comboBox.Enabled, "Failed at: Verify 'Combo Box' is Didabled failed");
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        /// <summary>
        /// Verify Combo Box is Disabled
        /// </summary>
        /// <param name="sWindow">Window</param>
        /// <param name="sComboBox">Combo Box ID</param>
        /// <param name="sLogInfo">Log test Setp</param>
        public static void verifyComboListEnabled(Window sWindow, string sComboBox, string sLogInfo = null)
        {
            try
            {
                ComboBox comboBox;
                char caseSwitch = (char)Common.ObjectIdentifier[sComboBox];
                switch (caseSwitch)

                {
                    case '#':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByText(sComboBox));
                        break;
                    case '?':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByClassName(sComboBox));
                        break;
                    case '$':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                    default:
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                }

                sWindow.Focus();

                //ATEL.LogTestSteps("Verify 'Combo Box' is Enabled", 2);

                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Assert.IsTrue(comboBox.Enabled, "Failed at: Verify 'Combo Box' is Enabled failed");
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        /// <summary>
        /// Verify Selected Combo Option
        /// </summary>
        /// <param name="sWindow">Window</param>
        /// <param name="sComboBox">Combo Box declaration</param>
        /// <param name="sExpectedSelection">Expected Selection</param>
        /// <param name="sLogInfo">Add Log Test Step (Optional)</param>
        public static void ComboVerifySelectedOption(Window sWindow, string sComboBox, string sExpectedSelection, string sLogInfo = null)
        {
            try
            {
                ComboBox comboBox;
                char caseSwitch = (char)Common.ObjectIdentifier[sComboBox];
                switch (caseSwitch)

                {
                    case '#':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByText(sComboBox));
                        break;
                    case '?':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByClassName(sComboBox));
                        break;
                    case '$':
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                    default:
                        comboBox = sWindow.Get<ComboBox>(SearchCriteria.ByAutomationId(sComboBox));
                        break;
                }

                sWindow.Focus();

                Assert.IsTrue(comboBox.SelectedItemText.Equals(sExpectedSelection), "Combo -- '" + "<B>Expected Selection: </B>" + sExpectedSelection + " <B>Current Selection: </B>" + comboBox.SelectedItemText);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        ////******Start CheckBox Function*************////
        /// <summary>
        /// Function to Select Checkbox
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="checkBoxID">Checkbox ID</param>
        /// <param name="sCheckBoxLabel">Checkbox Name</param>
        public static void chkBoxSelect(Window sWindow, String checkBoxID, string sCheckBoxLabel)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[checkBoxID];// Convert.ToChar(sField.Substring(0, 1));
                CheckBox chkBoxID;

                switch (caseSwitch)
                {
                    case '#':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByText(checkBoxID));
                        break;
                    case '?':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByClassName(checkBoxID));
                        break;
                    case '$':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                    default:
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                }
                chkBoxID.Focus();
                //ATEL.LogTestSteps("Select '" + sCheckBoxLabel + " ' CheckBox.");
                chkBoxID.Select();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Function to UnSelect Checkbox
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="checkBoxID">Checkbox ID</param>
        /// <param name="sCheckBoxLabel">Checkbox Name</param>
        public static void chkBoxUnSelect(Window sWindow, String checkBoxID, string sCheckBoxLabel)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[checkBoxID];// Convert.ToChar(sField.Substring(0, 1));
                CheckBox chkBoxID;

                switch (caseSwitch)
                {
                    case '#':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByText(checkBoxID));
                        break;
                    case '?':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByClassName(checkBoxID));
                        break;
                    case '$':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                    default:
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                }
                chkBoxID.Focus();
                //ATEL.LogTestSteps("UnSelect '" + sCheckBoxLabel + " ' CheckBox.");
                chkBoxID.UnSelect();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }


        /// <summary>
        /// Function to check if checkbox is Checked 
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="checkBoxID"Checkbox ID></param>
        /// <param name="sCheckBoxLabel">Checkbox Name</param>
        public static void verifyCheckBoxChecked(Window sWindow, String checkBoxID, string sCheckBoxLabel)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[checkBoxID];// Convert.ToChar(sField.Substring(0, 1));
                CheckBox chkBoxID;

                switch (caseSwitch)
                {
                    case '#':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByText(checkBoxID));
                        break;
                    case '?':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByClassName(checkBoxID));
                        break;
                    case '$':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                    default:
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                }
                //ATEL.LogTestSteps("Verify Checkbox for '" + sCheckBoxLabel + "' is Unchecked", 1);
                Assert.IsTrue(chkBoxID.IsSelected == true, "Checkbox: " + sCheckBoxLabel + " Is not checked.", 1);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        ////******Start Radio Button Functions*************////
        /// <summary>
        /// Function to Verify Radio Button is Enabled
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="radioButtonID">Radio Button ID</param>
        /// <param name="sRadioButtonLabel">Radio Button Name</param>
        public static void radioButtonEnabled(Window sWindow, String radioButtonID, string sRadioButtonLabel, int iDepth = 0)
        {
            CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = iDepth;

            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[radioButtonID];// Convert.ToChar(sField.Substring(0, 1));
                RadioButton rdoButtonID;

                switch (caseSwitch)
                {
                    case '#':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByText(radioButtonID));
                        break;
                    case '?':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByClassName(radioButtonID));
                        break;
                    case '$':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                    default:
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                }
                //ATEL.LogTestSteps("Verify Radio Button '" + sRadioButtonLabel + " ' is Enabled", 1);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Assert.IsTrue(rdoButtonID.Enabled, "Radio Button " + sRadioButtonLabel + " is not enabled");

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
                CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

            }
            CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

        }

        /// <summary>
        /// Function to Verify Radio Button is Disabled
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="radioButtonID">Radio Button ID</param>
        /// <param name="sRadioButtonLabel">Radio Button Name</param>
        public static void radioButtonDisabled(Window sWindow, String radioButtonID, string sRadioButtonLabel, int iDepth = 0)
        {
            CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = iDepth;

            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[radioButtonID];// Convert.ToChar(sField.Substring(0, 1));
                RadioButton rdoButtonID;

                switch (caseSwitch)
                {
                    case '#':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByText(radioButtonID));
                        break;
                    case '?':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByClassName(radioButtonID));
                        break;
                    case '$':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                    default:
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                }
                //ATEL.LogTestSteps("Verify Radio Button '" + sRadioButtonLabel + " ' is Disabled", 1);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Assert.IsTrue(rdoButtonID.Enabled == false, "Radio Button " + sRadioButtonLabel + " is  Disabled");

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search
        }

        /// <summary>
        /// Function to Verify Radio Button is Selected
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="radioButtonID">Radio Button ID</param>
        /// <param name="sRadioButtonLabel">Radio Button Name</param>
        public static void radioButtonSelected(Window sWindow, String radioButtonID, string sRadioButtonLabel, int iDepth = 0)
        {
            CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = iDepth;
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[radioButtonID];// Convert.ToChar(sField.Substring(0, 1));
                RadioButton rdoButtonID;

                switch (caseSwitch)
                {
                    case '#':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByText(radioButtonID));
                        break;
                    case '?':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByClassName(radioButtonID));
                        break;
                    case '$':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                    default:
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                }
                //ATEL.LogTestSteps("Verify Radio Button '" + sRadioButtonLabel + " ' is Selected", 1);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Assert.IsTrue(rdoButtonID.IsSelected, "Radio Button " + sRadioButtonLabel + " is  Selected");

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Function to Verify Radio Button is Selected
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="radioButtonID">Radio Button ID</param>
        /// <param name="sRadioButtonLabel">Radio Button Name</param>
        public static void radioButtonUnSelected(Window sWindow, String radioButtonID, string sRadioButtonLabel)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[radioButtonID];// Convert.ToChar(sField.Substring(0, 1));
                RadioButton rdoButtonID;

                switch (caseSwitch)
                {
                    case '#':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByText(radioButtonID));
                        break;
                    case '?':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByClassName(radioButtonID));
                        break;
                    case '$':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                    default:
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                }
                //ATEL.LogTestSteps("Verify Radio Button '" + sRadioButtonLabel + " ' is Selected", 1);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Assert.IsTrue(rdoButtonID.IsSelected == false, "Radio Button " + sRadioButtonLabel + " is not Selected");

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

        }

        /// <summary>
        /// Function to Click Radio Button 
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="radioButtonID">Radio Button ID</param>
        /// <param name="sRadioButtonLabel">Radio Button Name </param>
        public static void radioButtonSelect(Window sWindow, String radioButtonID, int iDepth = 0)
        {
            CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = iDepth;

            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[radioButtonID];// Convert.ToChar(sField.Substring(0, 1));
                RadioButton rdoButtonID;

                switch (caseSwitch)
                {
                    case '#':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByText(radioButtonID));
                        break;
                    case '?':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByClassName(radioButtonID));
                        break;
                    case '$':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                    default:
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                }
                //ATEL.LogTestSteps("Select Radio Button'" + rdoButtonID.Name.ToString() + " '", 1);
                rdoButtonID.Select();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
                CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

            }
            CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

        }

        internal static void clickTableRowCheckBox(Window batchCentrifugationWindow, string v)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Function to check if checkbox is UnChecked 
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="checkBoxID">Checkbox ID</param>
        /// <param name="sCheckBoxLabel">Checkbox Text</param>
        public static void verifyCheckBoxUnChecked(Window sWindow, String checkBoxID, string sCheckBoxLabel)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[checkBoxID];// Convert.ToChar(sField.Substring(0, 1));
                CheckBox chkBoxID;

                switch (caseSwitch)
                {
                    case '#':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByText(checkBoxID));
                        break;
                    case '?':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByClassName(checkBoxID));
                        break;
                    case '$':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                    default:
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                }
                //ATEL.LogTestSteps("Verify " + sCheckBoxLabel + " is unchecked");
                Assert.IsFalse(chkBoxID.IsSelected, "Checkbox: " + sCheckBoxLabel + " Is checked currently. Expected value UnChecked", 1);

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        public static void verifyCheckBoxVisible(Window sWindow, String checkBoxID, string sCheckBoxLabel)
        {
            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[checkBoxID];// Convert.ToChar(sField.Substring(0, 1));
                CheckBox chkBoxID;

                switch (caseSwitch)
                {
                    case '#':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByText(checkBoxID));
                        break;
                    case '?':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByClassName(checkBoxID));
                        break;
                    case '$':
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                    default:
                        chkBoxID = sWindow.Get<CheckBox>(SearchCriteria.ByAutomationId(checkBoxID));
                        break;
                }
                sCheckBoxLabel = (Regex.Replace(sCheckBoxLabel.Trim(), @"\s+", " "));
                string sCurrentDescription = "";

                //ATEL.LogTestSteps("Verify Checkbox '" + sCheckBoxLabel + "' Exists.", 1);

                if (!chkBoxID.Name.Equals(""))
                {
                    sCurrentDescription = (Regex.Replace(chkBoxID.Name.Trim(), @"\s+", " "));
                    Assert.IsTrue(sCurrentDescription.Equals(sCheckBoxLabel), "Expected Value : " + sCheckBoxLabel + " Current Value: " + sCurrentDescription);
                }
                else if (!chkBoxID.Text.Equals(""))
                {
                    sCurrentDescription = (Regex.Replace(chkBoxID.Text.Trim(), @"\s+", " "));
                    Assert.IsTrue(sCurrentDescription.Equals(sCheckBoxLabel), "Expected Value : " + sCheckBoxLabel + " Current Value: " + sCurrentDescription);
                }
                else
                {
                    Assert.Fail("Both Name & Text is not set for " + chkBoxID);
                }

            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }


        /// <summary>
        /// Verify Dialog Box and close the popup window
        /// </summary>
        /// <param name="currentApplication">Provide Application</param>
        /// <param name="sDIalogBoxTitle">Expected Popup DialogBox Title</param>
        /// <param name="sExpectedWarningMsg">Expected Warning Message</param>
        /// <param name="sClickButtonTxt">Button text to click</param>
        public static void VerifyDialog(Application currentApplication, string sDIalogBoxTitle, String sExpectedMessage, String sClickButtonTxt)
        {
            try
            {
                Window messageBox = Retry.For(() => currentApplication.GetWindows().First(x => x.Title.Contains("" + sDIalogBoxTitle)), TimeSpan.FromSeconds(30));

                messageBox.WaitWhileBusy();
                Assert.IsNotNull(messageBox, "Verify Message box Exissts");
                // var sCurrentWarningMsg = messageBox.Get<Label>(SearchCriteria.Indexed(0));
                string sCurrenMsg = Regex.Replace(messageBox.Get<Label>(SearchCriteria.Indexed(0)).Name.Trim(), @"\s+", " ");
                string sExpectedMsg = Regex.Replace(sExpectedMessage.Trim(), @"\s+", " ");
                //ATEL.LogTestSteps("Expected Warning Msg :" + sExpectedMsg, 1);
                Assert.AreEqual(sExpectedMsg.Trim(), sCurrenMsg, "<B>Expected Msg: </B>" + sExpectedMsg + " <B>Current Msg: </B>" + sCurrenMsg);
                //ATEL.LogTestSteps("Click on '" + sClickButtonTxt + "' Button", 1);
                messageBox.Get<Button>(SearchCriteria.ByText(sClickButtonTxt)).Click();
                Common.Sleep(2);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

        /// <summary>
        /// Provide Old File Name with Path and New Filename with path to replace to
        /// </summary>
        /// <param name="sOldFileName">Old FileName with Path</param>
        /// <param name="sNewFileName">New FIle Name with Path</param>
        public static void RenameFile(string sOldFileName, string sNewFileName)
        {
            try
            {
                if (File.Exists(sNewFileName))
                    File.Delete(sNewFileName);
                File.Move(sOldFileName, sNewFileName);
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        
        ////******End CheckBox Function*************////

        /// <summary>
        /// Checkes if the provided Search String exits in the file
        /// </summary>
        /// <param name="fileName">File Name</param>
        /// <param name="searchString">Search String</param>
        /// <returns></returns>
        public static bool IsStringInFile(string fileName, string searchString)
        {
            //ATEL.LogTestSteps("Verify String " + searchString + " Exits in " + fileName);
            if (File.ReadAllText(fileName).Contains(searchString))
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        /// <summary>
        /// Checkes if the exists
        /// </summary>
        /// <param name="fileName">File Name</param>
        /// <param name="searchString">Search String</param>
        /// <returns></returns>
        public static bool IsFileExists(string fileName)
        {

            //ATEL.LogTestSteps("Verify File " + fileName, 1);
            if (File.Exists(fileName))
            {
                return true;
            }
            else
            {
                return false;
            }

        }


        public static void RemoveReadOnlyFlag(DirectoryInfo di)
        {
            foreach (DirectoryInfo dif in di.GetDirectories())
            {
                RemoveReadOnlyFlag(dif);
            }

            foreach (FileInfo fi in di.GetFiles())
            {
                fi.Attributes = FileAttributes.Normal;
            }

            di.Attributes = FileAttributes.Normal;
        }

        /// <summary>
        /// Verify Folder Exists
        /// </summary>
        /// <param name="path"></param>
        public static bool IsFolderExists(string path)
        {
            bool bFound = false;
            // ... Set to folder path we must ensure exists.
            try
            {
                // ... If the directory doesn't exist, create it.
                if (Directory.Exists(path))
                {
                    bFound = true;
                }
            }
            catch (Exception)
            {
                // Fail silently.
            }
            return bFound;
        }

        /// <summary>
        /// Replaces text with new text in the specified file name
        /// </summary>
        /// <param name="sFile"></param>
        /// <param name="sCurrentText"></param>
        /// <param name="sReplaceToText"></param>
        public static void replaceText(string sFile, string sCurrentText, string sReplaceToText)
        {
            try
            {

                int i = 0;
                StringBuilder newFile = new StringBuilder();
                string temp = "";

                string[] file = System.IO.File.ReadAllLines(sFile);

                foreach (string line in file)
                {
                    if (line.Contains(sCurrentText))
                    {
                        //ATEL.LogTestSteps("Replacing text " + sCurrentText + " with " + sReplaceToText);
                        temp = line.Replace(sCurrentText, sReplaceToText);
                        i++;
                        newFile.Append(temp + "\r\n");
                        continue;
                    }
                    newFile.Append(line + "\r\n");
                }
                System.IO.File.WriteAllText(sFile, newFile.ToString());

            }

            catch (Exception e)
            {
                Common.Exception(e);
            }

        }

        /// <summary>
        ///  replateLine 
        /// </summary>
        /// <param name="sFile"></param>
        /// <param name="sLineContains"></param>
        /// <param name="sReplaceLine"></param>
        public static void replaceLine(string sFile, string sLineContains, String sReplaceLine)
        {

            try
            {
                int i = 0;
                StringBuilder newFile = new StringBuilder();
                string temp = "";

                string[] file = System.IO.File.ReadAllLines(sFile);

                foreach (string line in file)
                {
                    if (line.Contains(sLineContains))
                    {
                        string templine = line.ToString();
                        temp = line.Replace(templine, sReplaceLine);
                        i++;
                        newFile.Append(temp + "\r\n");
                        continue;
                    }
                    newFile.Append(line + "\r\n");
                }
                System.IO.File.WriteAllText(sFile, newFile.ToString());

            }

            catch (Exception e)
            {
                Common.Exception(e);
            }

        }

        /// <summary>
        /// Empty file contents
        /// </summary>
        /// <param name="sFile">Provide file name with path</param>
        public static void EmptyFile(string sFile)
        {
            try
            {
                System.IO.File.WriteAllText(sFile, string.Empty);
            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }

        /// <summary>
        ///
        ///</summary>
        /// <param name="sReplaceText">Replace Text</param>
        /// <param name="sNewText">New Text to replace to</param>
        /// <param name="sFilename">Provide complete file Path</param>
        public static void updateCSVFile(string sReplaceText, string sNewText, string sFilename)
        {
            try
            {
                //ATEL.LogTestSteps("Replacing text: " + sReplaceText + " to: " + sNewText + " in file " + sFilename, 2);
                string sBarCodeScanFile = @sFilename;
                String strFile = File.ReadAllText(sBarCodeScanFile);
                strFile = strFile.Replace(sReplaceText, sNewText);
                File.WriteAllText(sBarCodeScanFile, strFile);
            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }

        /// <summary>
        /// Click on Menu
        /// </summary>
        /// <param name="sWindow">Window</param>
        /// <param name="sMenuID">Menu ID</param>
        /// <param name="sMenuItem">Menu Option to select</param>
        public static void menuClick(Window sWindow, string sMenuID, string sMenuItem)
        {
            try
            {

                var menu = sWindow.Get(SearchCriteria.ByAutomationId(sMenuID));
                menu.Focus();
                
                //ATEL.LogTestSteps("Click on '" + sMenuItem);
                char caseSwitch = (char)Common.ObjectIdentifier[sMenuItem];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        sWindow.Get(SearchCriteria.ByText(sMenuItem)).Click();
                        break;
                    case '?':
                        sWindow.Get(SearchCriteria.ByClassName(sMenuItem)).Click();
                        break;
                    case '$':
                        sWindow.Get(SearchCriteria.ByAutomationId(sMenuItem)).Click();
                        break;
                    default:
                        sWindow.Get(SearchCriteria.ByAutomationId(sMenuItem)).Click();
                        break;
                }

                Common.Sleep(Common.obj_st_GlobalTestSupport.iTabTime);

            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }

        /// <summary>
        /// Copy Directory & Sub Direcorites
        /// </summary>
        /// <param name="source">Source Directory Location</param>
        /// <param name="target">Target Directory Location</param>
        public static void CopyFilesRecursively(DirectoryInfo source, DirectoryInfo target)
        {
            foreach (DirectoryInfo dir in source.GetDirectories())
                CopyFilesRecursively(dir, target.CreateSubdirectory(dir.Name));
            foreach (FileInfo file in source.GetFiles())
                file.CopyTo(Path.Combine(target.FullName, file.Name), true);
        }

        /// <summary>
        /// Delete Specified File
        /// </summary>
        /// <param name="sFromLocation"></param>
        public static void deleteFile(string sFromLocation)
        {
            try
            {
                //ATEL.LogTestSteps("Deleting File from" + sFromLocation, 0);

                if (File.Exists(sFromLocation))
                {
                    // File.SetAttributes(sFromLocation, FileAttributes.Normal);
                    File.Delete(sFromLocation);
                }


            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }

        public static void deleteFolder(string sFromLocation)
        {
            try
            {
                //ATEL.LogTestSteps("Deleting Folder from" + sFromLocation, 0);
                File.SetAttributes(sFromLocation, File.GetAttributes(sFromLocation) & ~FileAttributes.ReadOnly);

                if (Directory.Exists(sFromLocation))
                {
                    Directory.Delete(sFromLocation, true);
                }
            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }


        /// <summary>
        ///Copy File to specified location 
        ///</summary>
        /// <param name="sReplaceText">Replace Text</param>
        /// <param name="sNewText">New Text to replace to</param>
        /// <param name="sFilename">Provide complete file Path</param>
        public static void copyFile(string sFromLocation, string sToLocation)
        {
            try
            {
                //ATEL.LogTestSteps("Copying File from" + sFromLocation + " to: " + sToLocation, 0);

                if (File.Exists(sToLocation))
                {
                    File.Delete(sToLocation);
                }
                File.Copy(sFromLocation, sToLocation);

            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }


        /// <summary>
        /// Copy folder contents to specified folder
        /// </summary>
        /// <param name="sCurrentFolder"></param>
        /// <param name="sNewFolder"></param>
        public static void copyFolder(string sCurrentFolder, string sNewFolder)
        {
            try
            {
                //ATEL.LogTestSteps("Copying File from: " + sCurrentFolder + " to: " + sNewFolder, -2);

                foreach (string newPath in Directory.GetFiles(sCurrentFolder, "*.*", SearchOption.AllDirectories))
                {
                    File.Copy(newPath, newPath.Replace(sCurrentFolder, sNewFolder), true);
                }
            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }


        public static void CopyAll(DirectoryInfo source, DirectoryInfo target, bool overwrite/*, bool readOnly = false*/)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                //if ((fi.Name.ToUpper() != "Do Not add files to this folder!.txt".ToUpper()) && (fi.Name != "placeholder".ToUpper()))
                //{
                var fullName = Path.Combine(target.FullName, fi.Name);
                if (File.Exists(fullName))
                {
                    if (overwrite == true)
                    {
                        File.SetAttributes(fullName, FileAttributes.Normal);
                        fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
                        //App.Logger.Debug(CultureInfo.CurrentCulture, $"Copied '{fi.Name}' to '{target.FullName}'");
                    }
                }
                else
                {
                    fi.CopyTo(fullName, true);
                    //App.Logger.Debug(CultureInfo.CurrentCulture, $"Copied '{fi.Name}' to '{target.FullName}'");
                }
                //if (readOnly)
                //{
                //    File.SetAttributes(fullName, FileAttributes.ReadOnly);
                //    //App.Logger.Debug(CultureInfo.CurrentCulture, $"Set '{fi.Name}' attributes to ReadOnly");
                //}
                //}
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                    target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir, overwrite);
            }
        }

        public static void DeleteRecursiveFolder(DirectoryInfo dirInfo)
        {
            if (dirInfo.Exists)
            {
                foreach (var subDir in dirInfo.GetDirectories())
                {
                    DeleteRecursiveFolder(subDir);
                }

                foreach (var file in dirInfo.GetFiles())
                {
                    file.Attributes = FileAttributes.Normal;
                    file.Delete();
                }

                dirInfo.Delete();
            }

        }
        /// <summary>
        /// Verify displayed date format is expected format.
        /// </summary>
        /// <param name="sCurrentDate">Provide currently displayed date </param>
        /// <param name="expectedFormat">Provide Expected Date Format</param>
        public static void verifyDateformatIsCorrect(String sCurrentDate, string expectedFormat)
        {
            try
            {
                string tempdate;
                DateTime date;
                date = Convert.ToDateTime(sCurrentDate);
                tempdate = date.ToString(expectedFormat);
                //ATEL.LogTestSteps("Verify Current Date: " + sCurrentDate + "is in '" + expectedFormat + "' Format");
                Assert.IsTrue(sCurrentDate.Equals(tempdate), "<B>Expected Format </B>" + tempdate + " <B>Current Format </B>" + sCurrentDate);
            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }

        /// <summary>
        ///Verify that application is currently  running
        /// </summary>
        /// <param name="ProcessName"></param>
        /// <param name="bExpected"></param>
        public static void checkApplicationIsRunning(string ProcessName)
        {
            try
            {
                Process[] processes = Process.GetProcessesByName(ProcessName);
                //ATEL.LogTestSteps("Verify that " + ProcessName + "is currently  running.");

                if (processes.Length == 0)
                {
                    Assert.Fail(ProcessName + "is not currently not running");
                }

            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }

        /// <summary>
        ///Verify that application is currently not running
        /// </summary>
        /// <param name="ProcessName"></param>
        /// <param name="bExpected"></param>
        public static void checkApplicationIsNotRunning(string ProcessName)
        {
            try
            {
                Process[] processes = Process.GetProcessesByName(ProcessName);

                //ATEL.LogTestSteps("Verify that " + ProcessName + "is  currently not running.");
                if (processes.Length > 0)
                {
                    Assert.Fail(ProcessName + "is not currently running");
                }

            }
            catch (Exception e)
            {
                Common.Exception(e);
            }
        }

        

        //############# Functions to Set or Get values from INI files #############
        /*public static void UpdateAnyIniFileValue(string sFileName, string sSection, string sField, string sFieldValue, string sFilePath = null)
        {
            if (sFilePath == null)
                sFilePath = Common.obj_st_Location.sLocalPath;
            string sLocation = sFilePath + sFileName;  // File Location

            Common.IniFile ini = new Common.IniFile(sLocation);
             ATEL.LogTestSteps("Update the file: '" + sLocation + "', section: '" + sSection + "', field: '" + sField + "' with value: '" + sFieldValue + "'.");
            Console.WriteLine("Update the file: '" + sLocation + "', section: '" + sSection + "', field: '" + sField + "' with value: '" + sFieldValue + "'.");
            ini.Write(sField, sFieldValue, sSection);
        }

        public static string GetAnyIniFileValue(string sFileName, string sSection, string sField, string sFilePath = null)
        {
            string sTemp;
            if (sFilePath == null)
                sFilePath = Common.obj_st_Location.sLocalPath;
            string sLocation = sFilePath + sFileName;  // File Location

            Common.IniFile ini = new Common.IniFile(sLocation);
            //ATEL.LogTestSteps("Get the value from field: '" + sField + "', section: '" + sSection + "', file: '" + sLocation + "'.");
            Console.WriteLine("Get the value from field: '" + sField + "', section: '" + sSection + "', file: '" + sLocation + "'.");
            return sTemp = ini.Read(sField, sSection);
        }*/

        public static string GetRandomString(int length)
        {
            Random r = new Random();
            string charPool = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvw xyz1234567890";
            StringBuilder rs = new StringBuilder();

            while (length > 0)
            {
                rs.Append(charPool[(int)(r.NextDouble() * charPool.Length)]);
                length--;
            }
            return rs.ToString();
        }

        public static Boolean verifyDateFormat(string sDateOutput, string sExpectedDateFormat)
        {

            Boolean bDateFormat = false;
            if (sDateOutput == sExpectedDateFormat)
                bDateFormat = true;
            else
                bDateFormat = false;
            return bDateFormat;
        }

        /// <summary>
        /// Determines if string array is sorted from A -> Z
        /// </summary>
        public static bool IsSortedAscending(List<string> list)
        {
            for (int i = 0; i < list.Count; i++)
            {
                //ATEL.LogTestSteps("1" + (list[i] + "2" + list[i + 1]));
                if (list[i].CompareTo(list[i + 1]) < 0) // If previous is bigger, return false
                {
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// Determines if string array is sorted from A -> Z
        /// </summary>
        public static bool IsSortedDescending(List<string> list)
        {
            for (int i = 1; i < list.Count; i++)
            {
                //ATEL.LogTestSteps("1" + (list[i - 1]) + "2" + list[i]);
                if (list[i - 1].CompareTo(list[i]) < 0) // If previous is bigger, return false
                {
                    return false;
                }
            }
            return true;
        }



        /// <summary>
        /// Function to verify Column Headers contain expected text in header also return column headers
        /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnHeaderName">Column Header Name</param>
        /// <param name="sExpectedColumnList"> List of Row for the specified column will be verified with expected list </param>
        /// <returns>Returns list of Row contents for the columns</returns>
        public static List<string> verifyTableColumnContents(Window sWindow, String sTableID, string sColumnHeaderName, List<String> expectedColumnList = null)
        {
            List<string> currentColumnList = new List<string>();
            int iKeyIndex = 0;
            try
            {
                sWindow.ReloadIfCached();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Table table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

                for (int i = 0; i < table.Header.Columns.Count; i++)
                {
                    if (table.Header.Columns[i].NameMatches(sColumnHeaderName))
                    {
                        iKeyIndex = i;
                        //iKeyIndex = iKeyIndex + 1;
                        break;
                    }
                }

                if (iKeyIndex == -1)
                {
                    Assert.Fail("Column Name " + sColumnHeaderName + " does not exist");
                }
                int icount = table.Rows.Count;


                for (int i = 0; i < table.Rows.Count; i++)
                {
                    currentColumnList.Add(table.Rows[i].Cells[iKeyIndex].Value.ToString());
                }
                //Compare Expected Header List Columns with Current List of Comumns
                //ATEL.LogTestSteps("Verify Column Contents for " + sColumnHeaderName);

                if (expectedColumnList != null)
                {
                    for (int i = 0; i < expectedColumnList.Count; i++)
                    {
                        //ATEL.LogTestSteps("Verify " + (i + 1) + " Row text is set to " + expectedColumnList[i].ToString(), 2);
                        Assert.IsTrue((expectedColumnList[i].ToString().Equals(currentColumnList[i].ToString())), "<B>Expected Row Text</B>" + expectedColumnList[i].ToString() + " <B>Current Row Text:</B>" + currentColumnList[i].ToString());
                    }
                }
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return currentColumnList;
        }


        /// <summary>
        /// Function to verify Column Headers contain expected text in header also return column headers
        /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnHeaderName">Column Header Name</param>
        /// <param name="sExpectedColumnList"> List of Row for the specified column will be verified with expected list </param>
        /// <returns>Returns list of Row contents for the columns</returns>
        public static List<string> VerifyTableHeaders(Window sWindow, String sTableID, List<String> expectedColumnList = null)
        {
            List<string> currentColumnList = new List<string>();
            try
            {
                sWindow.ReloadIfCached();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Table table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

                //ATEL.LogTestSteps("Verify Table Column Headers" + table.Name);
                for (int i = 0; i < table.Header.Columns.Count; i++)
                {
                    //ATEL.LogTestSteps("Verify " + (i + 1) + " Column text is set to " + expectedColumnList[i].ToString(), 2);
                    Assert.IsTrue(table.Header.Columns[i].NameMatches(expectedColumnList[i]));
                }
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return currentColumnList;
        }

        /// <summary>
        ///  Function to verify Column content based on  Column Hearders to check from & value to check from Column 
        ///  /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnHeaderName">Column Header Name</param>
        /// <param name="sColumnFromHeaderName">Column From Header Name to Verify</param>
        /// <param name="sColumnToHeaderName">Column To Header to Verify</param>
        /// <param name="sColumntoCheckFrom">Column Row Title to check</param>
        /// <param name="sExpectedToValue">Expected Value in the To Column Row</param>
        public static void verifyTableRowContent(Window sWindow, String sTableID, string sColumnFromHeaderName, string sColumnToHeaderName, string sColumntoCheckFrom, string sExpectedToValue)
        {
            List<string> currentFromColumnList = new List<string>();
            List<string> currentToColumnList = new List<string>();

            int iKeyFromColumnIndex = -1;
            int iKeyToColumnIndex = -1;
            int iKeyRowIndex = -1;
            string sCurrentValue;
            try
            {
                sWindow.ReloadIfCached();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Table table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                for (int i = 0; i < table.Header.Columns.Count; i++)
                {
                    if (table.Header.Columns[i].NameMatches(sColumnFromHeaderName))
                        iKeyFromColumnIndex = i;
                    if (table.Header.Columns[i].NameMatches(sColumnToHeaderName))
                        iKeyToColumnIndex = i;
                }

                if (iKeyFromColumnIndex == -1)
                    Assert.Fail("Column Name " + sColumnFromHeaderName + " does not exist");
                if (iKeyToColumnIndex == -1)
                    Assert.Fail("Column Name " + sColumnToHeaderName + " does not exist");

                int icount = table.Rows.Count;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    currentFromColumnList.Add(table.Rows[i].Cells[sColumnFromHeaderName].Value.ToString());
                    if (table.Rows[i].Cells[sColumnFromHeaderName].Value.ToString().Equals(sColumntoCheckFrom))
                    {
                        iKeyRowIndex = i;
                    }
                }
                if (iKeyRowIndex == -1)
                    Assert.Fail("Column " + sColumnFromHeaderName + " does not contain " + sColumntoCheckFrom + " does not exist");

                //ATEL.LogTestSteps("Verify that '" + sColumnToHeaderName + "' value for '" + sColumntoCheckFrom + "' is set to " + sExpectedToValue, 1);
                sCurrentValue = table.Rows[iKeyRowIndex].Cells[sColumnToHeaderName].Value.ToString();
                Assert.IsTrue(sCurrentValue.Equals(sExpectedToValue), "<B> Expected Value for " + sColumntoCheckFrom + " is </B> " + sExpectedToValue + "<B>Current Value is </B> " + sCurrentValue);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        /// <summary>
        ///  Double click on Specific Row based on  Column Hearders to check from & Colum hearder click location 
        ///  /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnHeaderName">Column Header Name</param>
        /// <param name="sColumnFromHeaderName">Column From Header Name to Verify</param>
        /// <param name="sColumnToHeaderName">Column To Header to Verify</param>
        /// <param name="sColumntoCheckFrom">Column Row Title to check</param>
        /// <param name="sExpectedToValue">Expected Value in the To Column Row</param>
        public static void doubleClickTableRow(Window sWindow, String sTableID, string sColumnFromHeaderName, string sColumnToHeaderName, string sAction, string sColumntoCheckFrom, string sExpectedToValue)
        {
            List<string> currentFromColumnList = new List<string>();
            List<string> currentToColumnList = new List<string>();

            int iKeyFromColumnIndex = -1;
            int iKeyToColumnIndex = -1;
            int iKeyRowIndex = -1;
            string sCurrentValue;
            Table table;
            try
            {
                sWindow.ReloadIfCached();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

                char caseSwitch = (char)Common.ObjectIdentifier[sTableID];// Convert.ToChar(sField.Substring(0, 1));
                switch (caseSwitch)
                {
                    case '#':
                        table = sWindow.Get<Table>(SearchCriteria.ByText(sTableID));
                        break;
                    case '?':
                        table = sWindow.Get<Table>(SearchCriteria.ByClassName(sTableID));
                        break;
                    case '$':
                        table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                        break;
                    default:
                        table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                        break;
                }
                Common.Sleep(Common.obj_st_GlobalTestSupport.iTabTime);
                for (int i = 0; i < table.Header.Columns.Count; i++)
                {
                    if (table.Header.Columns[i].NameMatches(sColumnFromHeaderName))
                        iKeyFromColumnIndex = i;
                    if (table.Header.Columns[i].NameMatches(sColumnToHeaderName))
                        iKeyToColumnIndex = i;
                }
                if (iKeyFromColumnIndex == -1)
                    Assert.Fail("Column Name " + sColumnFromHeaderName + " does not exist");
                if (iKeyToColumnIndex == -1)
                    Assert.Fail("Column Name " + sColumnToHeaderName + " does not exist");

                int icount = table.Rows.Count;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    currentFromColumnList.Add(table.Rows[i].Cells[sColumnFromHeaderName].Value.ToString());
                    if (table.Rows[i].Cells[sColumnFromHeaderName].Value.ToString().Equals(sColumntoCheckFrom))
                    {
                        iKeyRowIndex = i;
                    }
                }
                if (iKeyRowIndex == -1)
                    Assert.Fail("Column " + sColumnFromHeaderName + " does not contain " + sColumntoCheckFrom + " does not exist");

                //ATEL.LogTestSteps("Select  '" + sColumnToHeaderName + "' Option '" + sColumntoCheckFrom + "' is set to " + sExpectedToValue, 1);
                table.Rows[iKeyRowIndex].Cells[sColumnToHeaderName].DoubleClick();
                table.Refresh();
                sCurrentValue = table.Rows[iKeyRowIndex].Cells[sColumnToHeaderName].Value.ToString();
                Assert.IsTrue(sCurrentValue.Equals(sExpectedToValue), "<B> Expected Value for " + sColumntoCheckFrom + " is </B> " + sExpectedToValue + "<B>Current Value is </B> " + sCurrentValue);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        /// <summary>
        /// Set value in table grid based on column name and row  provided by user
        /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnName">Column Name</param>
        /// <param name="sRowNumber">Row to Enter value for the specified column</param>
        /// <param name="sSetValue">Value to enter</param>
        public static void setTableRowText(Window sWindow, string sTableID, string sColumnName, int sRowNumber, String sSetValue)
        {
            try
            {
                //Default is set to first column
                int columnIndex = -1;
                int Row = sRowNumber - 1;

                Table table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                //Get Index location for the column
                for (int i = 0; i < table.Header.Columns.Count; i++)
                {
                    if (table.Header.Columns[i].NameMatches(sColumnName))
                    {
                        columnIndex = i;
                        break;
                    }
                }
                if (columnIndex == -1)
                {
                    Assert.Fail("Column Name " + sColumnName + " does not exist");
                }
                else
                {
                    //ATEL.LogTestSteps("Set 'Row " + sRowNumber + "'for '" + sColumnName + "' to " + sSetValue, 1);
                    table.Rows[Row].Cells[columnIndex].Focus();
                    table.Rows[Row].Cells[columnIndex].Click();
                    Keyboard.Instance.Enter(sSetValue);
                    Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
                    Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                }
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        /// <summary>
        /// Set value in table grid based on column name and row  provided by user
        /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnName">Column Name</param>
        /// <param name="sRowNumber">Row to Enter value for the specified column</param>
        /// <param name="sSetValue">Value to enter</param>
        public static void clickTableRowCheckBox(Window sWindow, string sTableID, string sColumnName, int sRowNumber, Boolean bValue)
        {
            try
            {
                //Default is set to first column
                int columnIndex = -1;
                int Row = sRowNumber - 1;

                Table table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                //Get Index location for the column
                for (int i = 0; i < table.Header.Columns.Count; i++)
                {
                    if (table.Header.Columns[i].NameMatches(sColumnName))
                    {
                        columnIndex = i;
                        break;
                    }
                }
                if (columnIndex == -1)
                {
                    Assert.Fail("Column Name " + sColumnName + " does not exist");
                }
                else
                {
                    //    ATEL.LogTestSteps("Set 'Row " + sRowNumber + "'for '" + sColumnName + "' to " + sSetValue, 1);
                    table.Rows[Row].Cells[columnIndex].Focus();
                    table.Rows[Row].Cells[columnIndex].Value = true;
                }
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }

        }

        /// <summary>
        /// get value in table grid based on column name and row  provided by user
        /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnName">Column Name</param>
        /// <param name="sRowNumber">get row value for the specified column</param>
        public static string getTableRowText(Window sWindow, string sTableID, string sColumnName, int sRowNumber)
        {
            var sText = "";
            int iKeyIndex = -1;
            try
            {
                sWindow.ReloadIfCached();
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);


                Table table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);

                for (int i = 0; i <= table.Header.Columns.Count; i++)
                {
                    if (table.Header.Columns[i].NameMatches(sColumnName))
                    {
                        iKeyIndex = i;
                        //iKeyIndex = iKeyIndex + 1;
                        break;
                    }
                }

                if (iKeyIndex == -1)
                {
                    Assert.Fail("Column Name " + sColumnName + " does not exist");
                }

                sText = table.Rows[sRowNumber].Cells[iKeyIndex].Value.ToString();
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return sText;
        }

        /// <summary>
        /// Set value in table grid based on column name and row  provided by user
        /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnName">Column Name</param>
        /// <param name="sRowNumber">Row to Enter value for the specified column</param>
        /// <param name="sSetValue">Value to enter</param>
        public static string getTableRowCheckBoxSelection(Window sWindow, string sTableID, string sColumnName, int sRowNumber)
        {
            string svalue = null;
            try
            {

                //Default is set to first column
                int columnIndex = -1;
                int Row = sRowNumber - 1;

                Table table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                //Get Index location for the column
                for (int i = 0; i <= table.Header.Columns.Count; i++)
                {
                    if (table.Header.Columns[i].NameMatches(sColumnName))
                    {
                        columnIndex = i;
                        break;
                    }
                }
                if (columnIndex == -1)
                {
                    Assert.Fail("Column Name " + sColumnName + " does not exist");
                }
                else
                {
                    //    ATEL.LogTestSteps("Set 'Row " + sRowNumber + "'for '" + sColumnName + "' to " + sSetValue, 1);
                    table.Rows[Row].Cells[columnIndex].Focus();
                    svalue = table.Rows[Row].Cells[columnIndex].Value.ToString();
                }
            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return svalue;
        }


        /// <summary>
        /// Set value in table grid based on column name and row  provided by user
        /// </summary>
        /// <param name="sWindow">Window name</param>
        /// <param name="sTableID">Table Declaration</param>
        /// <param name="sColumnName">Column Name</param>
        /// <param name="sRowNumber">Row to Enter value for the specified column</param>
        /// <param name="sSetValue">Value to enter</param>
        public static string VerifyTableRowBGColor(Window sWindow, string sTableID, string sColumnName, int sRowNumber, string sExpectedRGBColor)
        {
            string sColor = "test";

            try
            {
                //Default is set to first column
                int columnIndex = -1;
                int Row = sRowNumber - 1;


                Table table = sWindow.Get<Table>(SearchCriteria.ByAutomationId(sTableID));
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                //Get Index location for the column
                for (int i = 0; i <= table.Header.Columns.Count; i++)
                {
                    if (table.Header.Columns[i].NameMatches(sColumnName))
                    {
                        columnIndex = i;
                        break;
                    }
                }
                if (columnIndex == -1)
                {
                    Assert.Fail("Column Name " + sColumnName + " does not exist");
                }
                else
                {
                    //ATEL.LogTestSteps("Get BGColor for 'Row " + sRowNumber + "'for '" + sColumnName, 1);
                    table.Rows[Row].Cells[columnIndex].Focus();
                    sColor = table.Rows[Row].Cells[columnIndex].BorderColor.ToString();
                    //ATEL.LogTestSteps("Verify Background Color is set to " + sColor + " for  Row " + Row + " in Column " + sColumnName);
                    Assert.IsTrue(sColor.Equals(sExpectedRGBColor), "Expected BGColor :" + sExpectedRGBColor + " Current BGColor :" + sColor);

                }

            }

            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
            return sColor;
        }


        /// <summary>
        /// Verify Button is visible
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="searchCriteria">Button text to check it's visible</param>
        /// <param name="sbuttonText">Button Text (optional parameter)</param>
        public static void verifyButtonIsVisible(Window sWindow, string searchCriteria, string sbuttonText = null)
        {
            try
            {
                Button button;
                //char caseSwitch = Common.identifier;// Convert.ToChar(sField.Substring(0, 1));

                char caseSwitch = (char)Common.ObjectIdentifier[searchCriteria];
                switch (caseSwitch)
                {
                    case '#':
                        button = sWindow.Get<Button>(SearchCriteria.ByText(searchCriteria));
                        break;
                    case '?':
                        button = sWindow.Get<Button>(SearchCriteria.ByClassName(searchCriteria));
                        break;
                    case '$':
                        button = sWindow.Get<Button>(SearchCriteria.ByAutomationId(searchCriteria));
                        break;
                    default:
                        button = sWindow.Get<Button>(SearchCriteria.ByAutomationId(searchCriteria));
                        break;
                }
                if (button != null)
                {
                    string buttontext;
                    if (sbuttonText != null)
                        buttontext = sbuttonText;
                    else
                        buttontext = button.Text;

                    //ATEL.LogTestSteps("Verify Button " + buttontext + " is visible", 2);

                    if (button.Visible)
                        //ATEL.LogTestSteps("Button" + buttontext + " is visible", 2);
                    Assert.IsTrue(button.Visible, "Button is not currently Visible");
                }
                else
                {
                    Assert.IsFalse(button == null, "Button is not found in the current window with the given search criteria");
                }


            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }



        /// <summary>
        /// Function to Verify Radio Button is visible
        /// </summary>
        /// <param name="sWindow">Window Name</param>
        /// <param name="radioButtonID">Radio Button ID</param>
        /// <param name="sRadioButtonLabel">Radio Button Name</param>
        public static void radioButtonIsVisible(Window sWindow, String radioButtonID, string sRadioButtonLabel, int iDepth = 0)
        {
            CoreAppXmlConfiguration.Instance.MaxElementSearchDepth = iDepth;

            try
            {
                char caseSwitch = (char)Common.ObjectIdentifier[radioButtonID];// Convert.ToChar(sField.Substring(0, 1));
                RadioButton rdoButtonID;

                switch (caseSwitch)
                {
                    case '#':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByText(radioButtonID));
                        break;
                    case '?':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByClassName(radioButtonID));
                        break;
                    case '$':
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                    default:
                        rdoButtonID = sWindow.Get<RadioButton>(SearchCriteria.ByAutomationId(radioButtonID));
                        break;
                }

                //ATEL.LogTestSteps("Verify Radio Button '" + sRadioButtonLabel + " ' is visible", 1);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iwaitTime);
                Assert.IsTrue(rdoButtonID.Visible, "Radio Button " + sRadioButtonLabel + " is not visible");


            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
                CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

            }
            CoreAppXmlConfiguration.Instance.RawElementBasedSearch = false; // white starts using regular search

        }

        public static void SetSpinner(Window sWindow, String sSpinnerID, String sSetText, int iIndex = 0)
        {
            try
            {
                //string sCurrentText = "";
                char caseSwitch = (char)Common.ObjectIdentifier[sSpinnerID];
                sWindow.ReloadIfCached();
                var sSpinner = sWindow.Get(SearchCriteria.ByAutomationId(sSpinnerID).AndIndex(iIndex));
                sSpinner.Focus();

                sSpinner.DoubleClick();
                sSpinner.Focus();
                sSpinner.DoubleClick();
                Common.Sleep(2);
                Keyboard.Instance.Enter(sSetText);
                Common.Sleep(Common.obj_st_GlobalTestSupport.iTime);
                Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);
            }
            catch (Exception e)
            {
                Assert.Fail(e.Message);
            }
        }

       

        }
    }


