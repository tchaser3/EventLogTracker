/* Title:           Automated Production Reports Class
 * Date:            1-27-20
 * Author           Terry Holmes
 * 
 * Description:     This is used to do the automated reports */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NewEventLogDLL;
using DepartmentDLL;
using EmployeeProjectAssignmentDLL;
using DateSearchDLL;
using ProjectProductivityReportsDLL;

namespace EventLogTracker
{
    class AutomatedProductionReportsClass
    {
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        DepartmentClass TheDepartMentClass = new DepartmentClass();
        SendEmailClass TheSendEmailClass = new SendEmailClass();
        EmployeeProjectAssignmentClass TheEmployeeProjectAssignmentClass = new EmployeeProjectAssignmentClass();
        ProjectProductivityReportsClass TheProjectProductivityReportsClass = new ProjectProductivityReportsClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();

        FindSortedDepartmentDataSet TheFindSortedDepartmentDataSet = new FindSortedDepartmentDataSet();
        FindDepartmentProductionEmailByDepartmentIDDataSet TheFindDepartmentProductionEmailByDepartmentIDDataSet = new FindDepartmentProductionEmailByDepartmentIDDataSet();
        FindActiveDepartmentProductionEmailProjectsByDepartmentIDDataSet TheFindActiveDepartmentProductionEmailProjectsByDepartmentIDDataSet = new FindActiveDepartmentProductionEmailProjectsByDepartmentIDDataSet();
        FindPrivateProjectProductivityDateRangeDataSet TheFindPrivateProjectProductivityDateRangeDataSet = new FindPrivateProjectProductivityDateRangeDataSet();
        TotalProductivityDataSet TheTotalProductivityDataSet = new TotalProductivityDataSet();

        public void RunAutomatedProductionReports()
        {
            int intDepartmentCounter;
            int intDepartmentNumberOfRecords;
            int intProjectCounter;
            int intProjectNumberOfRecords;
            int intProductivityCounter;
            int intProductivityNumberOfRecords;
            int intEmployeeCounter;
            int intEmployeeNumberOfRecords;
            int intDepartmentID;
            DateTime datStartDate;
            DateTime datEndDate;
            string strProjectSuffix;
            string strFileName;
            string strPath = "\\\\bjc\\shares\\Documents\\ProductivityReports\\";
            bool blnFatalError = false;
            string strHeader = "Project Productivity Report";
            string strMessage;
            string strEmailAddress;

            try
            {
                datEndDate = DateTime.Now;
                datEndDate = TheDateSearchClass.RemoveTime(datEndDate);
                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 720);

                TheFindSortedDepartmentDataSet = TheDepartMentClass.FindSortedDepartment();

                intDepartmentNumberOfRecords = TheFindSortedDepartmentDataSet.FindSortedDepartment.Rows.Count - 1;

                for(intDepartmentCounter = 0; intDepartmentCounter <= intDepartmentNumberOfRecords; intDepartmentCounter++)
                {
                    intDepartmentID = TheFindSortedDepartmentDataSet.FindSortedDepartment[intDepartmentCounter].DepartmentID;
                    strFileName = TheFindSortedDepartmentDataSet.FindSortedDepartment[intDepartmentCounter].Department + Convert.ToString(DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + ".xlsx");

                    TheFindActiveDepartmentProductionEmailProjectsByDepartmentIDDataSet = TheDepartMentClass.FindActiveDepartmentProductionEmailProjectsByDepartmentID(intDepartmentID);

                    intProjectNumberOfRecords = TheFindActiveDepartmentProductionEmailProjectsByDepartmentIDDataSet.FindActiveDepartmentProductionEmailProjectsByDepartmentID.Rows.Count - 1;

                    TheTotalProductivityDataSet.totalproductivity.Rows.Clear();

                    if(intProjectNumberOfRecords > -1)
                    {
                        for(intProjectCounter = 0; intProjectCounter <= intProjectNumberOfRecords; intProjectCounter++)
                        {
                            strProjectSuffix = TheFindActiveDepartmentProductionEmailProjectsByDepartmentIDDataSet.FindActiveDepartmentProductionEmailProjectsByDepartmentID[intProjectCounter].ProjectSuffix;

                            TheFindPrivateProjectProductivityDateRangeDataSet = TheProjectProductivityReportsClass.FindPrivateProjectProductivityDateRange(strProjectSuffix, strProjectSuffix, datStartDate, datEndDate);

                            intProductivityNumberOfRecords = TheFindPrivateProjectProductivityDateRangeDataSet.FindPrivateProjectProductivityDateRange.Rows.Count - 1;

                            if(intProductivityNumberOfRecords > -1)
                            {
                                for(intProductivityCounter = 0; intProductivityCounter <= intProductivityNumberOfRecords; intProductivityCounter++)
                                {
                                    TotalProductivityDataSet.totalproductivityRow NewProductivityRow = TheTotalProductivityDataSet.totalproductivity.NewtotalproductivityRow();

                                    NewProductivityRow.AssignedProjectID = TheFindPrivateProjectProductivityDateRangeDataSet.FindPrivateProjectProductivityDateRange[intProductivityCounter].AssignedProjectID;
                                    NewProductivityRow.ProjectName = TheFindPrivateProjectProductivityDateRangeDataSet.FindPrivateProjectProductivityDateRange[intProductivityCounter].ProjectName;
                                    NewProductivityRow.TotalHours = TheFindPrivateProjectProductivityDateRangeDataSet.FindPrivateProjectProductivityDateRange[intProductivityCounter].TotalHours;

                                    TheTotalProductivityDataSet.totalproductivity.Rows.Add(NewProductivityRow);
                                }                                
                            }
                        }

                        blnFatalError = ExportToExcel(strFileName, strPath);

                        if (blnFatalError == true)
                            throw new Exception();

                        TheFindDepartmentProductionEmailByDepartmentIDDataSet = TheDepartMentClass.FindDepartmentProductionEmailByDepartmentID(intDepartmentID);

                        intEmployeeNumberOfRecords = TheFindDepartmentProductionEmailByDepartmentIDDataSet.FindDepartmentProductionEmailByDepartmentID.Rows.Count - 1;

                        intProductivityNumberOfRecords = TheTotalProductivityDataSet.totalproductivity.Rows.Count - 1;

                        strMessage = "<h1>Project Productivity Report</h1>";
                        strMessage += "<h1>An Excel Copy of the Report Can Be Found At " + strPath + "</h1>";
                        strMessage += "<table>";
                        strMessage += "<tr>";
                        strMessage += "<td><b>Assigned PProject ID</b></td>";
                        strMessage += "<td><b>Project Name</b></td>";
                        strMessage += "<td><b>Total Hours</b></td>";
                        strMessage += "</tr>";
                        strMessage += "<p>               </p>";

                        for(intProductivityCounter = 0; intProductivityCounter <= intProductivityNumberOfRecords; intProductivityCounter++)
                        {
                            //adding items to message
                            strMessage += "<tr>";
                            strMessage += "<td>" + TheTotalProductivityDataSet.totalproductivity[intProductivityCounter].AssignedProjectID + "</td>";
                            strMessage += "<td>" + TheTotalProductivityDataSet.totalproductivity[intProductivityCounter].ProjectName + "</td>";
                            strMessage += "<td>" + Convert.ToString(TheTotalProductivityDataSet.totalproductivity[intProductivityCounter].TotalHours) + "</td>";
                            strMessage += "</tr>";
                        }

                        for(intEmployeeCounter = 0; intEmployeeCounter <= intEmployeeNumberOfRecords; intEmployeeCounter++)
                        {
                            strEmailAddress = TheFindDepartmentProductionEmailByDepartmentIDDataSet.FindDepartmentProductionEmailByDepartmentID[intEmployeeCounter].EmailAddress;

                            blnFatalError = !(TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage));

                            if (blnFatalError == true)
                                throw new Exception();
                        }
                    }
                }

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automated Production Reports Class // Run Automated Production Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
            
        }
        private bool ExportToExcel(string strFileName, string strPath)
        {
            bool blnFatalError = false;

            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheTotalProductivityDataSet.totalproductivity.Rows.Count;
                intColumnNumberOfRecords = TheTotalProductivityDataSet.totalproductivity.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheTotalProductivityDataSet.totalproductivity.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheTotalProductivityDataSet.totalproductivity.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                workbook.SaveAs(strPath + strFileName);

            }
            catch (System.Exception ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automated Production Reports // Export To Excel " + ex.Message);

                TheMessagesClass.InformationMessage(ex.ToString());

                blnFatalError = true;
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }

            return blnFatalError;
        }
    }
}
