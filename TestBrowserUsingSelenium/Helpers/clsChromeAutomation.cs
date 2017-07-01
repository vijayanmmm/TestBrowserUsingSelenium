using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using System.Data;

namespace TestBrowserUsingSelenium.Helpers
{
    class clsChromeAutomation
    {
        Helpers.clsHelpers m_objHelper = new Helpers.clsHelpers();

        /// <summary>
        /// This function is to wait until an alert display, then accept the same alert, may click okay, yes. This function will wait for 50 seconds/based on the input parameter to display the alert if not then it will return false        
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strAlertTitle"> Title of the alert window</param>
        /// <param name="p_strAlertContentContains">Alert window text, this will be used for contains match</param>
        /// <returns></returns>
        public bool combas_AcceptAlertWindow(ChromeDriver p_chromeDriver, string p_strAlertContentContains, int intWaitingSeconds = 50)
        {
            //Wait till popup comes - wait for 50 seconds/based on the input
            WebDriverWait webWait = new WebDriverWait(p_chromeDriver, new TimeSpan(0, 0, intWaitingSeconds));

            try
            {
                webWait.Until(ExpectedConditions.AlertIsPresent());
            }
            catch (WebDriverTimeoutException) //Alert window not found
            {
                return false;
            }

            //Check whether the alert is present if not then show error
            try
            {
                //Check the alert title and content is same as the parameter
                if (p_chromeDriver.SwitchTo().Alert().Text.Contains(p_strAlertContentContains) == true)
                {
                    //Accept the alert  //Equal to click Ok /Okay / Yes
                    p_chromeDriver.SwitchTo().Alert().Accept();
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (NoAlertPresentException)
            {
                //return false - no alert found in 300 seconds
                return false;
            }
        }

        /// <summary>
        /// This function will convert the data of one table column into list
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="intCon1ColumnNo"></param>
        /// <param name="p_strTableXpath"></param>
        /// <param name="p_strTableID"></param>
        /// <returns></returns>
        public List<string> chromeGetTableSingleColumnValues(ChromeDriver p_chromeDriver, int intCon1ColumnNo, string p_strTableXpath, string p_strTableID = "")
        {
            List<string> lstColumnValues = new List<string>();

            IWebElement tblData = null;

            try
            {
                if (p_strTableXpath != "")
                    tblData = p_chromeDriver.FindElementByXPath(p_strTableXpath);
                else
                    tblData = p_chromeDriver.FindElementById(p_strTableID);
            }
            catch (Exception)
            {
                return lstColumnValues;
            }

            ReadOnlyCollection<IWebElement> rwEvents = tblData.FindElements(By.TagName("tr"));
            if (rwEvents != null)
            {
                lstColumnValues = new List<string>();
                string strCellValue;
                ReadOnlyCollection<IWebElement> cellsEvents = null;
                //Select each item in the table            
                for (int i = 0; i < rwEvents.Count; i++)
                {
                    cellsEvents = rwEvents[i].FindElements(By.TagName("td"));
                    if (cellsEvents.Count == 0) continue;
                    strCellValue = cellsEvents[intCon1ColumnNo].Text;
                    if (strCellValue == null) strCellValue = "";
                    lstColumnValues.Add(strCellValue);
                }
            }


            return lstColumnValues;
        }

        /// <summary>
        /// This function will return one row from the table for the matched two conditions
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strTableXpath"></param>
        /// <param name="intCon1ColumnNo"></param>
        /// <param name="strCon1MatchValue"></param>
        /// <param name="intCon2ColumnNo"></param>
        /// <param name="strCon2MatchValue"></param>
        /// <returns></returns>
        public ReadOnlyCollection<IWebElement> chromeGetTableRow(ChromeDriver p_chromeDriver, string p_strTableXpath, int intCon1ColumnNo, string strCon1MatchValue, int intCon2ColumnNo, string strCon2MatchValue)
        {
            IWebElement tblEvents = p_chromeDriver.FindElementByXPath(p_strTableXpath);
            ReadOnlyCollection<IWebElement> rwEvents = tblEvents.FindElements(By.TagName("tr"));
            if (rwEvents == null)
            {
                return null;
            }

            ReadOnlyCollection<IWebElement> cellsEvents = null;
            //Select each item in the table            
            for (int i = 0; i < rwEvents.Count; i++)
            {
                cellsEvents = rwEvents[i].FindElements(By.TagName("td"));
                if (cellsEvents == null) continue;
                if (cellsEvents.Count == 0) continue;
                //Check whether its RP Route plan or not.
                if (cellsEvents[intCon1ColumnNo].Text.ToLower() == strCon1MatchValue.ToLower() && cellsEvents[intCon2ColumnNo].Text.ToLower() == strCon2MatchValue.ToLower())
                {
                    return cellsEvents;
                }
            }

            return null;
        }

        /// <summary>
        /// This function will return the datatable form the input
        /// HeaderTableXPath is needed - if the header is in the different table
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strTableXpath"></param>
        /// <param name="p_strHeaderTableXPath"></param>
        /// <param name="strArrColumns"></param>
        /// <returns></returns>
        public DataTable chromeGetTable(ChromeDriver p_chromeDriver, string p_strTableXpath, string p_strHeaderTableXPath = "", string[] strArrColumns = null)
        {
            DataTable tblResult = new DataTable();
            //Maximize the window to get the tables clearly
            try { p_chromeDriver.Manage().Window.Maximize(); }
            catch { };

            int intFirstRow = 1;
            try
            {
                IWebElement tblData = p_chromeDriver.FindElementByXPath(p_strTableXpath);
                ReadOnlyCollection<IWebElement> rwData = tblData.FindElements(By.TagName("tr"));
                if (rwData == null)
                {
                    return null;
                }

                ReadOnlyCollection<IWebElement> headerCells = null;
                ReadOnlyCollection<IWebElement> cellsEvents = null;

                List<string> lstColumns = null;

                //Convert Array into list.....
                if (strArrColumns != null) lstColumns = strArrColumns.ToList<string>();

                //Find the header row
                if (p_strHeaderTableXPath == "")
                {
                    headerCells = rwData[0].FindElements(By.TagName("td"));
                }
                else
                {
                    IWebElement tblHeader = p_chromeDriver.FindElementByXPath(p_strHeaderTableXPath);
                    headerCells = tblHeader.FindElements(By.TagName("td"));
                    intFirstRow = 0;
                }

                //Add columns
                for (int i = 0; i < headerCells.Count; i++)
                {
                    if (lstColumns == null || lstColumns.Contains(headerCells[i].Text.Trim()) == true)
                        tblResult.Columns.Add(headerCells[i].Text.Trim(), typeof(string));
                }

                tblResult.Columns.Add("RowDblClickValue", typeof(string));

                //Add the values
                //Select each item in the table        
                for (int i = intFirstRow; i < rwData.Count; i++)
                {
                    DataRow drData = tblResult.NewRow();

                    //Store the attribute value
                    string strKeyValue = rwData[i].GetAttribute("ondblclick");
                    if (strKeyValue != null && strKeyValue != "" && strKeyValue.IndexOf('\'') != 0)
                    {
                        strKeyValue = strKeyValue.Split(new char[] { '\'' })[1];
                        drData["RowDblClickValue"] = strKeyValue;
                    }

                    //Check the columns
                    cellsEvents = rwData[i].FindElements(By.TagName("td"));
                    if (cellsEvents != null || cellsEvents.Count > 0)
                    {
                        for (int j = 0; j < cellsEvents.Count; j++)
                        {
                            if (lstColumns == null || lstColumns.Contains(headerCells[j].Text) == true)
                                drData[headerCells[j].Text.Trim()] = cellsEvents[j].Text.Trim();
                        }
                    }

                    tblResult.Rows.Add(drData);
                }

            }
            catch (Exception e)
            {
                e = e;
                throw;
            }

            return tblResult;
        }

        /// <summary>
        /// Set text to a text box
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strControlID"></param>
        /// <param name="p_strText"></param>
        public void chromeSetText(ChromeDriver p_chromeDriver, string p_strControlID, string p_strText)
        {
            try
            {
                p_chromeDriver.FindElementById(p_strControlID).Clear();
                p_chromeDriver.FindElementById(p_strControlID).SendKeys(p_strText);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// To get text from a text box
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strControlID"></param>
        /// <returns></returns>
        public string chromeGetText(ChromeDriver p_chromeDriver, string p_strControlID)
        {
            IWebElement iwebEle;
            try
            {
                iwebEle = p_chromeDriver.FindElementById(p_strControlID);
                return iwebEle.GetAttribute("Value");
            }
            catch (Exception)
            {
                throw;
            }

        }

        /// <summary>
        /// Set date , this is specific to one scenario where the date field is seperated 5 boxes, day, month, year, hour, miniutes
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_dtValue"></param>
        /// <param name="p_strHours"></param>
        /// <param name="p_strMinutes"></param>
        /// <param name="p_strYearControlID"></param>
        /// <param name="p_strMonthControlID"></param>
        /// <param name="p_strDayControlID"></param>
        /// <param name="p_strHoursControlID"></param>
        /// <param name="p_strMinutesControlID"></param>
        public void chromeSetDate(ChromeDriver p_chromeDriver, DateTime p_dtValue, string p_strHours, string p_strMinutes, string p_strYearControlID, string p_strMonthControlID, string p_strDayControlID, string p_strHoursControlID, string p_strMinutesControlID)
        {
            chromeSetText(p_chromeDriver, p_strYearControlID, p_dtValue.Year.ToString());
            chromeSetText(p_chromeDriver, p_strMonthControlID, m_objHelper.stringMakeTwoDigitString(p_dtValue.Month.ToString()));
            chromeSetText(p_chromeDriver, p_strDayControlID, m_objHelper.stringMakeTwoDigitString(p_dtValue.Day.ToString()));
            chromeSetText(p_chromeDriver, p_strHoursControlID, p_strHours);
            chromeSetText(p_chromeDriver, p_strMinutesControlID, p_strMinutes);
        }

        /// <summary>
        /// To get date , this is specific to one scenario where the date field is seperated 5 boxes, day, month, year, hour, miniutes
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strYearControlID"></param>
        /// <param name="p_strMonthControlID"></param>
        /// <param name="p_strDayControlID"></param>
        /// <param name="p_strHoursControlID"></param>
        /// <param name="p_strMinutesControlID"></param>
        /// <returns></returns>
        public DateTime chromeGetDate(ChromeDriver p_chromeDriver, string p_strYearControlID, string p_strMonthControlID, string p_strDayControlID, string p_strHoursControlID, string p_strMinutesControlID)
        {
            //Get all Values, then combine, then convert into date then return....
            DateTime dtReturnValue;
            string strDay, strMonth, strYear, strHour, strMinute;
            strYear = chromeGetText(p_chromeDriver, p_strYearControlID);
            strMonth = chromeGetText(p_chromeDriver, p_strMonthControlID);
            strDay = chromeGetText(p_chromeDriver, p_strDayControlID);

            strHour = chromeGetText(p_chromeDriver, p_strHoursControlID);
            strMinute = chromeGetText(p_chromeDriver, p_strMinutesControlID);

            dtReturnValue = Convert.ToDateTime(strYear + "/" + strMonth + "/" + strDay + " " + strHour + ":" + strMinute);
            return dtReturnValue;
        }

        /// <summary>
        /// This function is to find the element , if the chrome driver throws the error then null will be returned
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strElementID"></param>
        /// <returns></returns>
        public IWebElement chromeFindElementByID(ChromeDriver p_chromeDriver, string p_strElementID)
        {
            try
            {
                return p_chromeDriver.FindElementById(p_strElementID);
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Combox box to select a text
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strControlID"></param>
        /// <param name="p_strText"></param>
        public void chromeSelectElement_SelectText(ChromeDriver p_chromeDriver, string p_strControlID, string p_strText)
        {
            SelectElement selectControl = new SelectElement(p_chromeDriver.FindElementById(p_strControlID));
            //For Sea
            if (selectControl.SelectedOption.Text != p_strText)
                selectControl.SelectByText(p_strText);
        }

        /// <summary>
        /// Open the chrome browser - important- all existing chrome browser must be closed, otherwise it will not work
        /// Make sure you have the "ChromeDriver.exe" is available in the application executable location(in the same folder where the application exe resides)
        /// </summary>
        /// <param name="p_clnCloseOpenedChromDriver"></param>
        /// <returns></returns>
        public ChromeDriver openChrome(bool p_clnCloseOpenedChromDriver = true)
        {
            ChromeOptions choptions = new ChromeOptions();
            //choptions.BinaryLocation = "C:\\Program Files\\Google\\Chrome\\Application\\Chrom.exe";
            choptions.AddArgument("--test-type");
            choptions.AddArgument("--disable-plugins");
            choptions.AddArgument("--disable-extensions");
            //choptions.AddArgument("--enable-automation");
            choptions.AddArgument("--no-sandbox");
            choptions.AddUserProfilePreference("credentials_enable_service", false);
            choptions.AddUserProfilePreference("profile.password_manager_enabled", false);
            choptions.AddArgument("--start-maximized");

            //choptions.AddArgument("--allow-external-pages");
            //choptions.AddArgument("--allow-running-insecure-content");
            //choptions.AddArgument("--new-window");
            //choptions.AddArguments("--enable-strict-powerful-feature-restrictions");
            //choptions.AddUserProfilePreference("profile.default_content_setting_values.images", 0);


            if (p_clnCloseOpenedChromDriver == true)
            {
                //Close the previous chrome driver is it have opened            
                foreach (System.Diagnostics.Process proTem in System.Diagnostics.Process.GetProcessesByName("chromedriver"))
                {
                    proTem.CloseMainWindow();
                }
            }

            //Open the new chrome
            return (new ChromeDriver(Environment.CurrentDirectory, choptions, new TimeSpan(0, 5, 0)));
        }

        /// <summary>
        /// This function closes the google chrome and googlchrome driver
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <returns></returns>
        public bool closeChrome(ChromeDriver p_chromeDriver)
        {
            //Exit the Chrome Browser 
            p_chromeDriver.Close();
            //Quit the chrome driver
            p_chromeDriver.Quit();
            return true;
        }

        /// <summary>
        /// This function is to select the item in the combo box
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="strControlID"></param>
        /// <param name="strSelectText"></param>
        public void chromeSelectItem(ChromeDriver p_chromeDriver, string strControlID, string strSelectText)
        {
            SelectElement SelectElement = new SelectElement(p_chromeDriver.FindElementById(strControlID));
            //Select the option
            if (SelectElement.SelectedOption.Text != strSelectText)
                SelectElement.SelectByText(strSelectText);

            //Validation
            if (SelectElement.SelectedOption.Text != strSelectText)
            {
                throw new Exception("Validation failed for selecting the text:" + strSelectText + " for Control: " + strControlID);
            }
        }

        /// <summary>
        /// This function search a specific row in a table and returns the matched row
        /// </summary>
        /// <param name="p_chromeDriver"></param>
        /// <param name="p_strTableXpath"></param>
        /// <param name="p_strColumn1Name"></param>
        /// <param name="p_strCon1MatchValue"></param>
        /// <param name="p_strColumn2Name"></param>
        /// <param name="strCon2MatchValue"></param>
        /// <returns>Return type is ReadOnlyCollection-Add System.Collections.ObjectModel in using statement </returns>
        public IWebElement getTableRow(ChromeDriver p_chromeDriver, string p_strTableXpath, string p_strColumn1Name, string p_strCon1MatchValue, string p_strColumn2Name, string strCon2MatchValue)
        {
            IWebElement tblEvents = p_chromeDriver.FindElementByXPath(p_strTableXpath);
            ReadOnlyCollection<IWebElement> rwEvents = tblEvents.FindElements(By.TagName("tr"));
            int intColumn1Index = -1;
            int intColumn2Index = -1;
            if (rwEvents == null)
            {
                return null;
            }

            ReadOnlyCollection<IWebElement> cellsEvents = null;
            //Match the title and get the column index
            cellsEvents = rwEvents[0].FindElements(By.TagName("td"));
            for (int i = 0; i < cellsEvents.Count; i++)
            {
                if (cellsEvents[i].Text.Trim().ToLower() == p_strColumn1Name.ToLower())
                    intColumn1Index = i;
                else if (cellsEvents[i].Text.Trim().ToLower() == p_strColumn2Name.ToLower())
                    intColumn2Index = i;
            }

            //Check both index has been found, if not found then return null
            if (intColumn1Index == -1 || intColumn2Index == -1)
                return null;

            //Select each item in the table
            for (int i = 1; i < rwEvents.Count; i++)
            {
                cellsEvents = rwEvents[i].FindElements(By.TagName("td"));
                if (cellsEvents == null) continue;
                if (cellsEvents.Count == 0) continue;

                if (cellsEvents[intColumn1Index].Text.Trim().ToLower() == p_strCon1MatchValue.ToLower() && cellsEvents[intColumn2Index].Text.Trim().ToLower() == strCon2MatchValue.ToLower())
                {
                    return rwEvents[i];
                }
            }

            return null;
        }

        /// <summary>
        /// Get html table in the data table format
        /// </summary>
        /// <param name="p_htmlTable">table element</param>
        /// <param name="strArrColumns">specify the column name which only need to be retrieved, if dont speicfy then it will return all columns in the table</param>
        /// <returns></returns>
        public DataTable HTMLGetTable(HtmlAgilityPack.HtmlNode p_htmlTable, string[] strArrColumns = null)
        {
            DataTable tblResult = new DataTable();
            try
            {
                //Get the rows of the table
                HtmlAgilityPack.HtmlNodeCollection htmlRows = p_htmlTable.SelectNodes("tr");

                HtmlAgilityPack.HtmlNodeCollection headerCells = null;
                HtmlAgilityPack.HtmlNodeCollection cellsEvents = null;

                List<string> lstColumns = null;

                //Convert Array into list.....
                if (strArrColumns != null) lstColumns = strArrColumns.ToList<string>();

                //Add columns
                headerCells = htmlRows[0].SelectNodes("td");
                for (int i = 0; i < headerCells.Count; i++)
                {
                    if (lstColumns == null || lstColumns.Contains(headerCells[i].InnerText.Trim()) == true)
                        tblResult.Columns.Add(headerCells[i].InnerText, typeof(string));
                }

                //Add the values
                //Select each item in the table            
                for (int i = 1; i < htmlRows.Count; i++)
                {
                    DataRow drData = tblResult.NewRow();
                    //Check the columns
                    cellsEvents = htmlRows[i].SelectNodes("td");
                    if (cellsEvents != null || cellsEvents.Count > 0)
                    {
                        for (int j = 0; j < cellsEvents.Count; j++)
                        {
                            if (lstColumns == null || lstColumns.Contains(headerCells[j].InnerText) == true)
                                drData[headerCells[j].InnerText] = cellsEvents[j].InnerText.Trim();
                        }
                    }

                    tblResult.Rows.Add(drData);
                }

            }
            catch (Exception e)
            {
                e = e;
                throw;
            }

            return tblResult;
        }

    }
}
