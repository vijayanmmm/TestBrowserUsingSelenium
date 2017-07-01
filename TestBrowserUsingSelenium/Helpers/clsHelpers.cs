using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Microsoft.Office.Interop.Excel;

//This requires MicroSoft Access to be installed in the system, if its not installed then it will not work
//It requires only if you are going to use the functions in the region ExcelToDatTableORDataSet_UsingLinqToExcel - otherwise disable it
//using  LinqToExcel; 

using System.Threading;
using System.Runtime.InteropServices;

//its required only if you used the functions in region outlook and Browserlogin, otherwise comment it, so no need to add reference
using objAutoItx = AutoIt.AutoItX; 

//using System.Windows.Forms;


namespace TestBrowserUsingSelenium.Helpers
{
    class clsHelpers
    {
        #region ExcelToDatTableORDataSet_UsingLinqToExcel
            //public DataSet excelTodataSetUsingLinqToExcel(string strFilePath)
        //{
        //    DataSet dsStoreData = new DataSet();
        //    int i;
        //    var varExcel = new ExcelQueryFactory(strFilePath);

        //    //Open the excel and store it in the database            
        //    foreach (string strSheetName in varExcel.GetWorksheetNames())
        //    {
        //        System.Data.DataTable dt = new System.Data.DataTable(strSheetName);

        //        //Add the column names
        //        string[] varColumn = varExcel.GetColumnNames(strSheetName).ToArray();

        //        for (i = 0; i < varColumn.GetLength(0); i++)
        //        {
        //            if (varColumn[i].Contains("Priority") == true || varColumn[i].Contains("Record_ID") == true)
        //            {
        //                dt.Columns.Add(varColumn[i], typeof(int));
        //            }
        //            else
        //            {
        //                dt.Columns.Add(varColumn[i], typeof(String));
        //            }
        //        }

        //        //Copy the data
        //        var tem = from table in varExcel.Worksheet(strSheetName) select table;
        //        foreach (var item in tem)
        //        {
        //            //Validation to ignore the blank records
        //            if (item.Count < 1 || item[0] == null || item[0] == "") 
        //                continue;

        //            DataRow dr = dt.NewRow();
        //            for (i = 0; i < item.Count; i++)
        //            {
        //                //Check for null values, if null value then continue to next item
        //                if (item.Count < 1)
        //                    continue;

        //                //Convert into int, if the column data type is int
        //                if (dt.Columns[i].DataType == typeof(int)) 
        //                    dr[i] =  Convert.ToInt32(item[i]);
        //                else
        //                    dr[i] = item[i]; //Directly store the string value
        //            }

        //            //Add the new row in the data table
        //            dt.Rows.Add(dr);
        //        }

        //        dsStoreData.Tables.Add(dt);
        //        dt.Dispose();
        //        dt = null;
        //    }


        //    ////Testing

        //    //Query the SOB rules
        //    //DataRow[] drResult = dsStoreData.Tables["SOB_Rules"].Select("Team = 'All' or Team like 'LATAM'");
        //    //DataRow[] drResult1 = dsStoreData.Tables["SOB_Rules"].Select("(Team = 'All' or Team like 'LATAM') and (Team <> 'IIII')");

        //    //DataTable dtTable = dsStoreData.Tables["SOB_Rules"].Select("'VAR OPPORTUNITY' like '%' + Team + '%'").CopyToDataTable();

        //    return dsStoreData;
        //}        
        #endregion

        #region DataTable_To_Excel
            public string dataTableToExcel(System.Data.DataTable p_dtData)
        {
            if (p_dtData == null) return "Data table is empty";
            Application appExcel = new Application();
            _Worksheet ws;

            //Set screen updating false
            appExcel.ScreenUpdating = false;

            //Add new excel workbook
            ws = appExcel.Workbooks.Add().Sheets[1];

            //set the automatic calculation is manual
            appExcel.Calculation = XlCalculation.xlCalculationManual;

            //Update the column title
            for (int i = 0; i < p_dtData.Columns.Count; i++)
            {
                ws.Cells[1, i + 1].Value = p_dtData.Columns[i].ColumnName;
            }

            //Loop through data table rows and columns
            for (int i = 0; i < p_dtData.Rows.Count; i++)
            {
                for (int j = 0; j < p_dtData.Columns.Count; j++)
                {
                    ws.Cells[i + 2, j + 1].Value = p_dtData.Rows[i][j];
                }
            }

            //Set the back to true for the below two
            //set the automatic calculation is manual and screenupdating false to improve the performance
            appExcel.Calculation = XlCalculation.xlCalculationAutomatic;
            appExcel.ScreenUpdating = true;
            //Make workbook Visible
            appExcel.Visible = true;

            return "Success";
        }
        #endregion
        
        #region Excel_To_DataSet_Or_DataTable_UsingExcelObject
            //Note the below one will work without the Access reference, purely using excel object
            //And with fully memory handling, After the data captured from excel, the object will be released completely from memory, you can check in the task manager to confirm(ofcourse spend lot of time on search in stackoverflow and other websites, thanks to buddies who helped me on this

            /// <summary>
            /// This function will import the excel to datatable
            /// It will take the first sheet in the excel
            /// the columns is defined the consecutive values in the first row, if any blank it means the columns count will end eventhough it has columns after a blank column
            /// The row values for first column should have values for all, if no value then tool will ignore the rest.
            /// </summary>
            /// <param name="strFilePath"></param>
            /// <returns></returns>
            public System.Data.DataTable excelToDataTable_UsingObjectArray_WithMemoryHandling(string strFilePath)
            {
                Application appExcel = null;
                _Workbook wb = null;
                _Worksheet ws = null;
                Range rngCell = null;
                System.Data.DataTable dt = null;
                int i, j;
                int intRowCount;
                int intColumnCount;
                int intUTNIndex = -1;
                string strRowCellValue;

                try
                {
                    //Open the excel and store it in the database
                    appExcel = new Application();
                    wb = appExcel.Workbooks.Open(strFilePath, ReadOnly: true);
                    //Take the first sheet
                    ws = wb.Sheets[1];
                    string strColumn = null;
                    dt = new System.Data.DataTable(ws.Name);
                    intRowCount = 0;
                    intColumnCount = 0;

                    //Add the values to the datatable - use the Array to get the values - to increase the speed
                    //Add the column Names
                    for (i = 1; i <= 16000; i++)
                    {
                        rngCell = ws.Cells[1, i];
                        strColumn = rngCell.Value;

                        //If blank column then exit the loop
                        if (strColumn == null || strColumn.ToString().Trim() == "")
                        {
                            Marshal.FinalReleaseComObject(rngCell);
                            break;
                        }

                        //to get the column count
                        intColumnCount = intColumnCount + 1;

                        if (strColumn.Contains("Priority") == true || strColumn.Contains("Record_ID") == true)
                        {
                            dt.Columns.Add(ws.Cells[1, i].Value, typeof(int));
                        }
                        else
                        {
                            try
                            {
                                dt.Columns.Add(strColumn, typeof(string));
                            }
                            catch (DuplicateNameException)
                            {
                                dt.Columns.Add(strColumn + "1", typeof(string));
                            }

                            if (strColumn.Contains("UTN") == true)
                                intUTNIndex = i;
                        }

                        Marshal.FinalReleaseComObject(rngCell);
                    }
                    dt.AcceptChanges();


                    //Get the row count
                    for (i = 1; i <= 1000000; i++)
                    {
                        rngCell = ws.Cells[i, 1];
                        strRowCellValue = rngCell.Text;

                        //If blank column then skip
                        if (strRowCellValue == null || strRowCellValue.ToString().Trim() == "")
                        {
                            Marshal.FinalReleaseComObject(rngCell);
                            break;
                        }

                        intRowCount = intRowCount + 1;
                        Marshal.FinalReleaseComObject(rngCell);
                    }

                    //Check the column and rows count
                    if (dt.Columns.Count == 0) return null;

                    //Get the used columns count
                    Range rngBegin, rngEnd;
                    rngBegin = ws.Cells[1, 1];
                    rngEnd = ws.Cells[intRowCount, intColumnCount];

                    rngCell = ws.Range[rngBegin, rngEnd];
                    object[,] objArrayValues = (object[,])rngCell.Value;
                    Marshal.FinalReleaseComObject(rngBegin);
                    Marshal.FinalReleaseComObject(rngEnd);
                    Marshal.FinalReleaseComObject(rngCell);

                    for (i = 2; i <= objArrayValues.GetLength(0); i++)
                    {
                        DataRow dr = dt.NewRow();
                        for (j = 1; j <= objArrayValues.GetLength(1); j++)
                        {
                            try
                            {
                                objArrayValues[i, j] = objArrayValues[i, j];
                            }
                            catch (IndexOutOfRangeException)
                            {
                                break;
                            }

                            if (objArrayValues[i, j] == null || objArrayValues[i, j].ToString() == "")
                            {
                                //Ignore the blank cell alues
                                continue;
                            }

                            if (objArrayValues[1, j].ToString().Contains("Priority") == true || objArrayValues[1, j].ToString().Contains("Record_ID") == true)
                            {
                                dr[j - 1] = Convert.ToInt32(objArrayValues[i, j].ToString());
                            }
                            else
                            {
                                dr[j - 1] = objArrayValues[i, j].ToString();
                            }
                        }

                        //Validate the row is blank, if so ignore the same
                        if (string.Join("", dr.ItemArray).ToString().Trim() != "")
                        {
                            dt.Rows.Add(dr);
                            dt.AcceptChanges();
                        }
                    }


                    //If the first column name is UTN, then get the XCP number - This number used to get the information from XCP page directly
                    if (intUTNIndex != -1 && dt.Columns.Contains("XCPNumber") == false)
                    {
                        dt.Columns.Add("XCPNumber", typeof(string));
                        //Get the index of the UTN
                        for (i = 2; i <= intRowCount; i++)
                        {
                            rngCell = ws.Cells[i, 1];
                            //Get the hyperlink
                            try
                            {
                                strRowCellValue = rngCell.Hyperlinks[1].Address;
                            }
                            catch (Exception)
                            {
                                continue;
                            }

                            //If blank column then skip
                            if (strRowCellValue == null || strRowCellValue.ToString().Trim() == "")
                            {
                                Marshal.FinalReleaseComObject(rngCell);
                                break;
                            }
                            strRowCellValue = strRowCellValue.Substring(strRowCellValue.IndexOf("fileid=") + 7);

                            dt.Rows[i - 2]["XCPNumber"] = strRowCellValue;
                            Marshal.FinalReleaseComObject(rngCell);
                        }
                        dt.AcceptChanges();
                    }

                    wb.Close(SaveChanges: false);
                    appExcel.Quit();
                }
                catch (Exception ex)
                {
                    ex = ex;
                    throw ex;
                }
                finally
                {
                    if (rngCell != null) Marshal.FinalReleaseComObject(rngCell);
                    if (ws != null) Marshal.FinalReleaseComObject(ws);
                    if (wb != null) Marshal.FinalReleaseComObject(wb);
                    if (appExcel != null) Marshal.FinalReleaseComObject(appExcel);
                }

                return dt;
            }

            /// <summary>
            /// This function helps to read the "Table" from all sheets - **import data should be in the table format(excel tables)
            /// Each sheet data should be entered in the form of table. no data outside table will be retrived
            /// Prerequiste-one excel sheet should have only one table, sheet name and the table name should be same
            /// </summary>
            /// <param name="strFilePath"></param>
            /// <returns></returns>
            public DataSet excelTableTodataSet_WithMemoryRelease(string strFilePath)
            {
                DataSet dsStoreData = new DataSet();
                Application appExcel = new Application();
                _Workbook wb = null;
                Workbooks wbs = null;
                Sheets wss = null;
                ListObjects LstObjs = null;
                _Worksheet ws = null;
                ListObject p_xlTable = null;

                try
                {
                    //Open the excel and store it in the database
                    wbs = appExcel.Workbooks;
                    wb = wbs.Open(strFilePath, ReadOnly: true);
                    wss = wb.Sheets;
                    ////Loop through all excel sheets                
                    for (int i = 1; i <= wss.Count; i++)
                    {
                        ws = wss[i];
                        LstObjs = ws.ListObjects;
                        for (int j = 1; j <= LstObjs.Count; j++)
                        {
                            p_xlTable = LstObjs[j];
                            dsStoreData.Tables.Add(convertExcelTableToDataTable(p_xlTable));
                            Marshal.FinalReleaseComObject(p_xlTable);
                        }
                        Marshal.FinalReleaseComObject(LstObjs);
                        Marshal.FinalReleaseComObject(ws);
                    }

                    //Close the excel                                                
                    wb.Close(SaveChanges: false);
                    appExcel.Application.Quit();
                }
                catch (Exception ex)
                {
                    ex = ex;
                    throw ex;
                }
                finally
                {
                    //Release memory            
                    if (wss != null) Marshal.FinalReleaseComObject(wss);
                    if (wb != null) Marshal.FinalReleaseComObject(wb);
                    if (wbs != null) Marshal.FinalReleaseComObject(wbs);
                    if (appExcel != null) Marshal.FinalReleaseComObject(appExcel);
                }

                return dsStoreData;
            }

            /// <summary>
            /// This function is convert the excel tables into the data table.. this is the supporting function for the above two functions
            /// </summary>
            /// <param name="p_xlTable"></param>
            /// <returns></returns>
            private System.Data.DataTable convertExcelTableToDataTable(ListObject p_xlTable)
            {
                int i, j;
                System.Data.DataTable dt = new System.Data.DataTable(p_xlTable.Name);
                Range rngTable = p_xlTable.Range;
                string strColumnName;

                //Use the object range to get the values
                //Add the values to the datatable - directly get the value from the excel cells                  
                object[,] objArrayValues = (object[,])rngTable.Value;


                //Add the columns to the database
                for (j = 1; j <= objArrayValues.GetLength(1); j++)
                {
                    strColumnName = objArrayValues[1, j].ToString();
                    if (strColumnName.Contains("Priority") == true || strColumnName.Contains("Record_ID") == true)
                    {
                        dt.Columns.Add(strColumnName, typeof(int));
                    }
                    else
                    {
                        dt.Columns.Add(strColumnName, typeof(string));
                    }
                }
                dt.AcceptChanges();


                for (i = 2; i <= objArrayValues.GetLength(0); i++)
                {
                    DataRow dr = dt.NewRow();
                    for (j = 1; j <= objArrayValues.GetLength(1); j++)
                    {
                        try
                        {
                            objArrayValues[i, j] = objArrayValues[i, j];
                        }
                        catch (IndexOutOfRangeException)
                        {
                            break;
                        }

                        if (objArrayValues[i, j] == null || objArrayValues[i, j].ToString() == "")
                        {
                            //Ignore the blank cell alues
                            continue;
                        }

                        if (objArrayValues[1, j].ToString().Contains("Priority") == true || objArrayValues[1, j].ToString().Contains("Record_ID") == true)
                        {
                            dr[j - 1] = Convert.ToInt32(objArrayValues[i, j].ToString());
                        }
                        else
                        {
                            dr[j - 1] = objArrayValues[i, j].ToString();
                        }
                    }

                    //Validate the row is blank, if so ignore the same
                    if (string.Join("", dr.ItemArray).ToString().Trim() != "")
                    {
                        dt.Rows.Add(dr);
                        dt.AcceptChanges();
                    }
                }

                Marshal.FinalReleaseComObject(rngTable);
                Marshal.FinalReleaseComObject(p_xlTable);

                //Return
                return dt;
            }

        #endregion

        #region DataTable

            /// <summary>
            /// This function will combine two different tables into one
            /// </summary>
            /// <param name="p_dtFirst"></param>
            /// <param name="p_dtSecond"></param>
            /// <returns></returns>
            public System.Data.DataTable combineTwoTables(System.Data.DataTable p_dtFirst, System.Data.DataTable p_dtSecond)
        {
            //Loop through second table columns and add the columns into the new table
            //Also add in the first table
            foreach (DataColumn dc in p_dtSecond.Columns)
            {
                if (p_dtFirst.Columns.Contains(dc.ColumnName) == false)
                    p_dtFirst.Columns.Add(dc.ColumnName, dc.DataType);
            }
            p_dtFirst.AcceptChanges();

            foreach (DataColumn dc in p_dtFirst.Columns)
            {
                if (p_dtSecond.Columns.Contains(dc.ColumnName) == false)
                    p_dtSecond.Columns.Add(dc.ColumnName, dc.DataType);
            }
            p_dtSecond.AcceptChanges();

            //Loop through second table and add all the values to the table one
            foreach (DataRow dr in p_dtSecond.Rows)
            {
                p_dtFirst.ImportRow(dr);
            }
            p_dtFirst.AcceptChanges();

            return p_dtFirst;
        }

        #endregion

        #region String
        public string stringGetAplhaNumeric(string p_strText)
        {
            string strReturnValue = "";
            byte[] byteArr = Encoding.ASCII.GetBytes(p_strText);
            if (p_strText == null || p_strText.Trim() == "") return "";

            for (int i = 0; i < byteArr.Length; i++)
            {
                if ((byteArr[i] >= 65 && byteArr[i] <= 90) || (byteArr[i] >= 97 && byteArr[i] <= 122) || (byteArr[i] >= 48 && byteArr[i] <= 57))
                {
                    strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                }
            }

            return strReturnValue;
        }

        public string stringGetAplha(string p_strText)
        {
            string strReturnValue = "";
            byte[] byteArr = Encoding.ASCII.GetBytes(p_strText);
            if (p_strText == null || p_strText.Trim() == "") return "";

            for (int i = 0; i < byteArr.Length; i++)
            {
                if ((byteArr[i] >= 65 && byteArr[i] <= 90) || (byteArr[i] >= 97 && byteArr[i] <= 122))
                {
                    strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                }
            }

            return strReturnValue;
        }

        public string stringGetNumeric(string p_strText)
        {
            string strReturnValue = "";
            byte[] byteArr = Encoding.ASCII.GetBytes(p_strText);
            if (p_strText == null || p_strText.Trim() == "") return "";

            for (int i = 0; i < byteArr.Length; i++)
            {
                if (byteArr[i] >= 48 && byteArr[i] <= 57)
                {
                    strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                }
            }

            return strReturnValue;
        }

        /// <summary>
        /// This method will check whether the base string contains the match string using 80% matching
        /// </summary>
        /// <param name="p_strBaseString"></param>
        /// <param name="p_strMatchString"></param>
        /// <returns>IT will return the matched string from the base string</returns>
        public string findContainsTextwith85PercentMatch(string p_strBaseString, string p_strMatchString)
        {
            string strReturnValue;
            List<byte> lstMatched = new List<byte>();

            byte[] byteBaseString = Encoding.ASCII.GetBytes(p_strBaseString);
            byte[] byteMatchString = Encoding.ASCII.GetBytes(p_strMatchString);

            int intFailedMatch = 0;
            int intMatchStringIndex = 0;
            int intAllowedFailure = p_strMatchString.Length - (p_strMatchString.Length * 90 / 100);
            if (p_strMatchString.Length < 6)
                intAllowedFailure = 0;


            //Match Each letters
            for (int i = 0; i < byteBaseString.Length; i++)
            {
                //quite if the predict that the match will not happen
                if ((intMatchStringIndex + (byteBaseString.Length - i)) < (byteMatchString.Length - intAllowedFailure))
                    return "";

                if (byteBaseString[i] == byteMatchString[intMatchStringIndex])
                {
                    if (intMatchStringIndex == 0) lstMatched.Clear();

                    lstMatched.Add(byteBaseString[i]);
                    intMatchStringIndex = intMatchStringIndex + 1;
                    //Check whether the matchstring already completed
                    if (byteMatchString.Length <= intMatchStringIndex)
                    {
                        //return the matched text
                        //Remove unncessary items, if the length is higher
                        if (lstMatched.Count > byteMatchString.Length && lstMatched[0] == lstMatched[1])
                            lstMatched.Remove(lstMatched[0]);
                        strReturnValue = ASCIIEncoding.ASCII.GetString(lstMatched.ToArray<byte>());
                        return strReturnValue;
                    }
                }
                else if (intMatchStringIndex > 0 && (intMatchStringIndex + 1 < byteMatchString.Length) && byteBaseString[i] == byteMatchString[intMatchStringIndex + 1])
                {
                    lstMatched.Add(byteBaseString[i]);
                    intMatchStringIndex = intMatchStringIndex + 2;
                    intFailedMatch = intFailedMatch + 1;

                    //Check if there is failed match or not..
                    if (intFailedMatch > intAllowedFailure)
                    {
                        lstMatched.Clear();
                        intFailedMatch = 0;
                        intMatchStringIndex = 0;
                    }
                    else
                    {
                        //Check whether the matchstring already completed
                        if (byteMatchString.Length <= (intMatchStringIndex))
                        {
                            //Remove unncessary items, if the length is higher
                            if (lstMatched.Count > byteMatchString.Length && lstMatched[0] == lstMatched[1])
                                lstMatched.Remove(lstMatched[0]);
                            //return the matched text
                            strReturnValue = ASCIIEncoding.ASCII.GetString(lstMatched.ToArray<byte>());
                            return strReturnValue;
                        }
                    }
                }
                else
                {
                    lstMatched.Add(byteBaseString[i]);
                    intFailedMatch = intFailedMatch + 1;
                    if (intFailedMatch > intAllowedFailure)
                    {
                        lstMatched.Clear();
                        intFailedMatch = 0;
                        intMatchStringIndex = 0;
                    }
                    else
                    {
                        //Check whether the matchstring already completed
                        if (byteMatchString.Length <= (intMatchStringIndex))
                        {
                            //Remove unncessary items, if the length is higher
                            if (lstMatched.Count > byteMatchString.Length && lstMatched[0] == lstMatched[1])
                                lstMatched.Remove(lstMatched[0]);
                            //return the matched text
                            strReturnValue = ASCIIEncoding.ASCII.GetString(lstMatched.ToArray<byte>());
                            return strReturnValue;
                        }
                    }
                }
            }

            return "";
        }

        public bool string_Compare90PercentMatch(string strBaseText, string strCompareText)
        {
            byte[] byteBaseString = Encoding.ASCII.GetBytes(strBaseText);
            byte[] byteCompareString = Encoding.ASCII.GetBytes(strCompareText);
            int intFailedMatch = 0;
            int intAllowedFailure = byteBaseString.Length - byteBaseString.Length * 85 / 100;

            if (byteBaseString.Length == byteCompareString.Length)
            {
                //Match Each letters
                for (int i = 0; i < byteBaseString.Length; i++)
                {
                    if (byteBaseString[i] != byteCompareString[i])
                        intFailedMatch = intFailedMatch + 1;
                }

                //Apply 85% percent match
                if (intFailedMatch <= intAllowedFailure)
                    return true;
                else
                    return false;
            }

            //Logic to compare neareast match.
            //take first "BaseString length and use the same with compare string to compare.....
            //Use this method only if 15% gap
            if (byteBaseString.Length > byteCompareString.Length && byteBaseString.Length < (byteCompareString.Length + intAllowedFailure))
            {
                int intGap = byteBaseString.Length - intAllowedFailure;
                if (strBaseText.Substring(0, intGap) == strCompareText.Substring(0, intGap))
                    return true;

                //Run through each letter then, ignore the no of allowed failure match
                for (int i = 0; i < byteCompareString.Length; i++)
                {
                    if (byteBaseString.Length < (i + intFailedMatch + 1))
                        return false;
                    if (byteCompareString[i] != byteBaseString[i + intFailedMatch])
                        intFailedMatch = intFailedMatch + 1;
                }
                //Compare how many letters failed, if its within allowed failure then return true
                //Apply 85% percent match
                if (intFailedMatch <= intAllowedFailure)
                    return true;
            }
            else if (byteBaseString.Length < byteCompareString.Length && (byteBaseString.Length + intAllowedFailure) > byteCompareString.Length)
            {
                int intGap = byteCompareString.Length - intAllowedFailure;
                if (strBaseText.Substring(0, intGap) == strCompareText.Substring(0, intGap))
                    return true;

                //Run through each letter then, ignore the no of allowed failure match
                for (int i = 0; i < byteBaseString.Length; i++)
                {
                    if (byteCompareString.Length < (i + intFailedMatch + 1))
                        return false;
                    if (byteBaseString[i] != byteCompareString[i + intFailedMatch])
                        intFailedMatch = intFailedMatch + 1;
                }
                //Compare how many letters failed, if its within allowed failure then return true
                //Apply 85% percent match
                if (intFailedMatch <= intAllowedFailure)
                    return true;
            }




            return false;
        }

        public string stringGetLastRow(string p_strText)
        {
            if (p_strText == null) return "";
            if (p_strText.IndexOf('\n') <= 0) return p_strText;

            string[] strArray = p_strText.Split('\n');
            return strArray[strArray.Length - 1];
        }

        public List<string> stringArrayToList(string[] strArrText)
        {
            List<string> lstReturnValue = new List<string>();
            foreach (string strText in strArrText)
            {
                if (strText.Trim() != "")
                    lstReturnValue.Add(strText);
            }

            return lstReturnValue;
        }

        /// <summary>
        /// This function will return two letter string, if it has only one letter then it will add "0" at the beginning.
        /// </summary>
        /// <param name="p_strText"></param>
        /// <returns></returns>
        public string stringMakeTwoDigitString(string p_strText)
        {
            if (p_strText == null) return "";
            if (p_strText.Length == 1) p_strText = "0" + p_strText;

            return p_strText;
        }

        private string string_HtmlTrim(string p_strText)
        {
            return p_strText.Replace("\t", "").Replace("\n", "").Replace("\r", "").Replace("&nbsp;", "").Replace("&amp;", "").Replace("&quot;", "").Replace("&gt;", "").Trim();
        }

        /// <summary>
        /// This function is to get the first two words in a text
        /// </summary>
        /// <param name="p_strText"></param>
        /// <returns></returns>
        public string string_GetFirstTwoWords(string p_strText)
        {
            //search the words
            string[] strTemArr = p_strText.Split(' ');
            if (strTemArr.Length > 2)
                return strTemArr[0] + " " + strTemArr[1];
            else
                return p_strText;
        }

        public string string_GetFirstWord(string p_strText)
        {
            //search the words
            string[] strTemArr = p_strText.Split(' ');
            if (strTemArr.Length > 2)
                return strTemArr[0];
            else
                return p_strText;
        }


        public string string_GetSecondWord(string p_strText)
        {
            //search the words
            string[] strTemArr = p_strText.Split(' ');
            if (strTemArr.Length > 1)
                return strTemArr[1];
            else
                return p_strText;
        }

        /// <summary>
        /// This function will keep only the alpha and numberic and space, others will be removed
        /// This function also seperates the string and word, if there is no gap inbetween
        /// </summary>
        /// <param name="p_strText"></param>
        /// <returns></returns>
        public string string_RemoveSpecialCharacters(string p_strText)
        {
            string strReturnValue = "";
            byte[] byteArr = Encoding.ASCII.GetBytes(p_strText);
            int intItemCode = -1; //0 for alpha 1 for number 2 for space or line feed
            if (p_strText == null || p_strText.Trim() == "") return "";

            for (int i = 0; i < byteArr.Length; i++)
            {
                if ((byteArr[i] >= 65 && byteArr[i] <= 90) || (byteArr[i] >= 97 && byteArr[i] <= 122)) //Alpha
                {
                    if (intItemCode == 1)
                        strReturnValue = strReturnValue + " " + Convert.ToChar(byteArr[i]);
                    else
                        strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                    intItemCode = 0;
                }
                else if ((byteArr[i] >= 48 && byteArr[i] <= 57)) //Numeric
                {
                    //Check the last letter is alpha if so then add one space
                    if (intItemCode == 0)
                        strReturnValue = strReturnValue + " " + Convert.ToChar(byteArr[i]);
                    else
                        strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                    intItemCode = 1;
                }
                else if (byteArr[i] == 32) //Space , line feed and comma
                {
                    if (intItemCode != 2) strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                    intItemCode = 2;
                }
                else if (byteArr[i] == 10) //Space , line feed and comma
                {
                    if (strReturnValue == "")
                        strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                    else
                    {
                        if (strReturnValue.Substring(strReturnValue.Length - 1) == " ")
                            strReturnValue = strReturnValue.Substring(0, strReturnValue.Length - 1) + Convert.ToChar(byteArr[i]);
                        else
                            strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                    }
                    intItemCode = 2;
                }
                else if (byteArr[i] == 44 || byteArr[i] == 46)
                {
                    if (intItemCode != 2) strReturnValue = strReturnValue + Convert.ToChar(32);
                    intItemCode = 2;
                }
            }

            return strReturnValue;
        }

        public string string_GetContinousNo(string p_strText, bool blnAcceptOneFailure = false)
        {
            string strReturnValue = "";
            bool blnOneFailed = false;
            byte[] byteArr = Encoding.ASCII.GetBytes(p_strText);
            if (p_strText == null || p_strText.Trim() == "") return "";

            for (int i = 0; i < byteArr.Length; i++)
            {
                if (byteArr[i] >= 48 && byteArr[i] <= 57)
                {
                    strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                }
                else
                {
                    if (strReturnValue.Length >= 4)
                        return strReturnValue;
                    else
                    {
                        if (strReturnValue.Length > 0 && blnAcceptOneFailure == true && blnOneFailed == false)
                            blnOneFailed = false;
                        else if (strReturnValue.Length > 0 && blnAcceptOneFailure == false)
                            return string_GetContinousNo(p_strText, true);
                        else
                            strReturnValue = "";
                    }
                }

            }

            return strReturnValue;
        }

        /// <summary>
        /// This function will provide the text, which match the text given and followed by few numbers
        /// The prequirement of this function is no special characters except Space
        /// </summary>
        /// <param name="p_strText"></param>
        /// <param name="p_strMatchText"></param>
        /// <returns></returns>
        public string string_GetContinousNumbersAfterText(string p_strText, string p_strMatchText)
        {
            string strReturnValue = "";
            string strBaseText;
            bool blnNumbersFound = false;

            strBaseText = p_strText.Trim();

            if (strBaseText == null || strBaseText.Trim() == "") return "";
            strBaseText = strBaseText.ToLower();
            if (strBaseText.Contains(p_strMatchText.ToLower()) == false) return "";

            //It means the text is matched..
            strReturnValue = p_strMatchText;
            byte[] byteArr = Encoding.ASCII.GetBytes(strBaseText.Substring(strBaseText.IndexOf(p_strMatchText) + p_strMatchText.Length));

            for (int i = 0; i < byteArr.Length; i++)
            {
                if (byteArr[i] >= 48 && byteArr[i] <= 57)
                {
                    strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                    blnNumbersFound = true;
                }
                else if (byteArr[i] == 32)
                {
                    strReturnValue = strReturnValue + Convert.ToChar(byteArr[i]);
                }
                else
                {
                    break;
                }
            }

            if (blnNumbersFound == true)
                return strReturnValue;
            else
                return "";
        }

        public bool string_IsBlank(object objText)
        {
            if (objText == null)
                return true;
            else if (objText.ToString().Trim() == "")
                return true;
            else
                return false;
        }

        /// <summary>
        /// This function will combine the multiple string into one, adding space inbetween
        /// </summary>
        /// <param name="strText1"></param>
        /// <param name="strText2"></param>
        /// <param name="strText3"></param>
        /// <param name="strText4"></param>
        /// <param name="strText5"></param>
        /// <param name="strText6"></param>
        /// <param name="strText7"></param>
        /// <param name="strText8"></param>
        /// <param name="strText9"></param>
        /// <param name="strText10"></param>
        /// <returns></returns>
        public string stringCombineWithSpace(string strDelimitor, string strText1, string strText2, string strText3 = "", string strText4 = "", string strText5 = "", string strText6 = "", string strText7 = "", string strText8 = "", string strText9 = "", string strText10 = "")
        {
            string strReturnValue = strText1 + " " + strText2;

            if (strText3 != "") strReturnValue = strReturnValue + strDelimitor + strText3;
            if (strText4 != "") strReturnValue = strReturnValue + strDelimitor + strText4;
            if (strText5 != "") strReturnValue = strReturnValue + strDelimitor + strText5;
            if (strText6 != "") strReturnValue = strReturnValue + strDelimitor + strText6;
            if (strText7 != "") strReturnValue = strReturnValue + strDelimitor + strText7;
            if (strText8 != "") strReturnValue = strReturnValue + strDelimitor + strText8;
            if (strText9 != "") strReturnValue = strReturnValue + strDelimitor + strText9;
            if (strText10 != "") strReturnValue = strReturnValue + strDelimitor + strText10;

            return strReturnValue;
        }

        /// <summary>
        /// This function will combine the array into one string with space as the delimitor
        /// </summary>
        /// <param name="p_strArrText"></param>
        /// <returns></returns>
        public string stringCombineWithSpace(string[] p_strArrText, string strDelimitor = " ")
        {
            string strReturnValue = "";

            foreach (string strText in p_strArrText)
            {
                if (strText.Trim() == "") continue;
                if (strReturnValue == "")
                    strReturnValue = strText;
                else
                    strReturnValue = strReturnValue + strDelimitor + strText;
            }

            return strReturnValue;
        }

        /// <summary>
        /// This function checks whether the object is empty or not, if empty then returns true, if not empty then returns false
        /// This is only for string
        /// </summary>
        /// <param name="p_obj"></param>
        /// <returns></returns>
        public bool IsEmpty(object p_obj)
        {
            if (p_obj == null) return true;
            if (p_obj.ToString().Trim() == "") return true;

            return false;
        }

        /// <summary>
        /// This function checks whether the object is empty or not, if empty then returns true, if not empty then returns false
        /// This is only for string
        /// </summary>
        /// <param name="p_obj"></param>
        /// <returns></returns>
        public bool IsDBNULL(object p_obj)
        {
            if (p_obj == DBNull.Value) return true;
            if (p_obj.ToString().Trim() == "") return true;

            return false;
        }

        #endregion

        #region Outlook
            System.Threading.Thread thOutlookClickAllow;
            /// <summary>
            /// These functions are in a case where you face outlook is asking permission to send email, each email
            /// you can include this class and activate it by calling ActiviateAutoClickAllow
            /// it will monitor whether the outlook warning message appears, if it appears it will autclick accept
            /// to stop this by calling DeActivateAutoClickAllow
            /// </summary>
            /// <returns></returns>
            public string outlook_ActiviateAutoClickAllow()
            {
                try
                {
                    //Start the outlook auto choose allow thread
                    thOutlookClickAllow = new Thread(new ThreadStart(outlook_AutoClickAllow));
                    thOutlookClickAllow.IsBackground = false;
                    thOutlookClickAllow.Start();
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }

                return "Success";
            }

            public string outlook_DeActivateAutoClickAllow()
            {
                try
                {
                    if (thOutlookClickAllow != null)
                    {
                        //Check whether the thread is alive if alive then abort it
                        if (thOutlookClickAllow.IsAlive)
                        {
                            thOutlookClickAllow.Abort();
                            Thread.Sleep(1000);
                        }
                        //Set the null value
                        thOutlookClickAllow = null;
                    }
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }

                return "Success";
            }

            private void outlook_AutoClickAllow()
            {
                string strWinTitle, StrWinContainsText, strWinContainsText_ToSendEmail, strtem;

                strWinTitle = "Microsoft Outlook";
                StrWinContainsText = "A program is trying to access e-mail address information stored in Outlook. If this is unexpected, click Deny and verify your antivirus software is up-to-date";
                strWinContainsText_ToSendEmail = "A program is trying to send an e-mail message on your behalf. If this is unexpected, click Deny and verify your antivirus software is up-to-date";

                try
                {

                    while (true)
                    {
                        //Wait for .3 seconds
                        Thread.Sleep(300);

                        //Check weather the outlook is showing the warning message or not.
                        //Set the Auto ITX mode to search full name of the window

                        while (objAutoItx.WinExists(strWinTitle, StrWinContainsText) == 1)
                        {
                            //MessageBox.Show("Window found, Click the Allow Access check box");

                            //1. Click the Allow Access Check Box
                            strtem = "-1";
                            while ((strtem == "-1" || strtem == "0" || strtem == "") && objAutoItx.ControlCommand(strWinTitle, StrWinContainsText, "ComboBox1", "IsEnabled", "") == "0")
                            {
                                //Enable the control then click the button
                                objAutoItx.WinActivate(strWinTitle, StrWinContainsText);
                                strtem = objAutoItx.ControlFocus(strWinTitle, StrWinContainsText, "Button3").ToString();
                                strtem = objAutoItx.ControlClick(strWinTitle, StrWinContainsText, "Button3").ToString();
                                Thread.Sleep(1000);
                            }

                            //2.1 Show the drop box and Select the string 10 minutes
                            strtem = "-1";
                            while ((strtem == "-1" || strtem == "0" || strtem == "") && objAutoItx.ControlGetText(strWinTitle, StrWinContainsText, "ComboBox1").ToString() != "10 minutes")
                            {
                                //Enable the control then click the button
                                objAutoItx.WinActivate(strWinTitle, StrWinContainsText);
                                strtem = objAutoItx.ControlFocus(strWinTitle, StrWinContainsText, "ComboBox1").ToString();
                                strtem = objAutoItx.ControlCommand(strWinTitle, StrWinContainsText, "ComboBox1", "SelectString", "10 minutes").ToString();
                                Thread.Sleep(500);
                            }

                            //3. Click the Allow Button
                            strtem = "-1";
                            while ((strtem == "-1" || strtem == "0" || strtem == "") && objAutoItx.WinExists(strWinTitle, StrWinContainsText) == 1)
                            {
                                //MessageBox.Show("Window found, Click the Allow Button double click");
                                //Enable the control then click the button
                                objAutoItx.WinActivate(strWinTitle, StrWinContainsText);
                                strtem = objAutoItx.ControlFocus(strWinTitle, StrWinContainsText, "Button4").ToString();
                                strtem = objAutoItx.ControlClick(strWinTitle, StrWinContainsText, "Button4", "left", 2).ToString();
                                objAutoItx.Send("{ENTER}");
                            }
                        }

                        //Warning message for the outlook email message
                        while (objAutoItx.WinExists(strWinTitle, strWinContainsText_ToSendEmail) == 1)
                        {
                            while (objAutoItx.ControlCommand(strWinTitle, strWinContainsText_ToSendEmail, "Button4", "IsEnabled", "") == "0")
                            {
                                Thread.Sleep(100);
                            }

                            strtem = "-1";
                            while ((strtem == "-1" || strtem == "0" || strtem == "") && objAutoItx.WinExists(strWinTitle, strWinContainsText_ToSendEmail) == 1)
                            {
                                //MessageBox.Show("Window found, Click the Allow Button double click");
                                //Enable the control then click the button                            
                                objAutoItx.WinActivate(strWinTitle, strWinContainsText_ToSendEmail);
                                strtem = objAutoItx.ControlFocus(strWinTitle, strWinContainsText_ToSendEmail, "Button4").ToString();
                                strtem = objAutoItx.ControlClick(strWinTitle, strWinContainsText_ToSendEmail, "Button4", "left", 2).ToString();
                                if (objAutoItx.WinExists(strWinTitle, strWinContainsText_ToSendEmail) == 1) objAutoItx.Send("{ENTER}");
                            }
                        }

                        //Looping 
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Error occured:" + ex.Message);
                }
            }
        #endregion

        #region BrowserLogin

            System.Threading.Thread  thBrowserLogin;
            /// <summary>
            /// These functions is to click ok in the chrome window
            /// Start it using ActiviateAutoClickALogin function with browser title
            /// Stop it using DeActivateAutoClickLogin function
            /// </summary>
            /// <param name="strBrowserTitle">The browser title for the web page</param>
            /// <returns></returns>
            public string BrowserLogin_ActiviateAutoClickALogin(string strBrowserTitle)
            {
                try
                {
                    //Start the outlook auto choose allow thread                    
                    thBrowserLogin = new Thread(new ParameterizedThreadStart(BrowserLogin_AutoClickLogin));
                    thBrowserLogin.IsBackground = false;
                    thBrowserLogin.Start(strBrowserTitle);
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }

                return "Success";
            }

            public string BrowserLogin_DeActivateAutoClickLogin()
            {
                try
                {
                    if (thBrowserLogin != null)
                    {
                        //Check whether the thread is alive if alive then abort it
                        if (thBrowserLogin.IsAlive)
                        {
                            thBrowserLogin.Abort();
                            Thread.Sleep(1000);
                        }
                        //Set the null value
                        thBrowserLogin = null;
                    }
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }

                return "Success";
            }

            private void BrowserLogin_AutoClickLogin(object p_strBrowserTitle)
            {
                
                try
                {
                    while (true)
                    {
                        //Looping 
                        //Wait for .3 seconds
                        Thread.Sleep(300);

                        //Warning message for the outlook email message
                        while (objAutoItx.WinExists(p_strBrowserTitle.ToString(), "") == 1)
                        {
                            Thread.Sleep(2000);
                            IntPtr winhanle = objAutoItx.WinGetHandle(p_strBrowserTitle.ToString(), "");                            
                            objAutoItx.WinActivate(winhanle);
                            objAutoItx.Send("{ENTER}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    //System.Windows.Forms.MessageBox.Show("Error occured:" + ex.Message);
                }
            }

        #endregion

        /// <summary>
        /// This function is to display the open file dialog box,after user chose the file, it will return the file path as output
        /// </summary>
        /// <param name="strFileType">Excel only- currently it supports(Coded) only excel</param>
        /// <returns>File path</returns>
        public string openFile(string strFileType = "Excel")
            {
                string strPath = "";

                //Get the database file
                if (strFileType == "Excel")
                {

                    System.Windows.Forms.OpenFileDialog ObjOpenExcel = new System.Windows.Forms.OpenFileDialog();
                    ObjOpenExcel.Filter = "Excel files (*.xlsx)|*.xls*";
                    if (ObjOpenExcel.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        strPath = "";
                    }
                    else
                    {
                        strPath = ObjOpenExcel.FileName;
                    }
                }

                return strPath;
            }

        /// <summary>
        /// This function is to update log or some text into a file
        /// Ensure the path is correct and the file exist to avoid issues
        /// </summary>
        /// <param name="strFilePath"></param>
        /// <param name="strText"></param>
        /// <returns></returns>
        public bool updateLog(string strFilePath, string strText)
        {
            //Assign to the stream writter
            System.IO.StreamWriter m_fswLogFile = new System.IO.StreamWriter(strFilePath, true);
            //Update in the file          
            m_fswLogFile.WriteLine(strText);
            m_fswLogFile.Close();
            return true;
        }        

        //Destructors
        ~clsHelpers()
        {
            if (thOutlookClickAllow != null) outlook_DeActivateAutoClickAllow();
            //if (thOutlookClickAllow != null) objAutoItx = null;            
        }       

    }
}
