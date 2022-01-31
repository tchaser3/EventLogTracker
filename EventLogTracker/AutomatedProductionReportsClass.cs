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
using DesignProductivityDLL;
using EmployeeDateEntryDLL.FindEmployeeDataEntryByDateRangeDataSetTableAdapters;
using Microsoft.Office.Core;

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
        DesignProductivityClass TheDesignProductivityClass = new DesignProductivityClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();

        FindAllEmployeeProductionOverAWeekDataSet TheFindAllEmployeesProductionOverAWeekDataSet = new FindAllEmployeeProductionOverAWeekDataSet();
        FindAllDesignEmployeeProductivityOverAWeekDataSet TheFindAllDesignEmployeeProductivityOverAWeekDataSet = new FindAllDesignEmployeeProductivityOverAWeekDataSet();
        EmployeeProductivityDataSet TheEmployeeProductivityDataSet = new EmployeeProductivityDataSet();
        CompleteProjectProductivityDataSet TheCompleteProjectProductivityDataSet = new CompleteProjectProductivityDataSet();
        FindProductivityManagersForEmailDataSet TheFindProductivityManagersForEmailDataSet = new FindProductivityManagersForEmailDataSet();

        //setting global variables
        int gintProjectCounter;
        int gintProjectNumberOfRecords;

        public void RunAutomatedProductionReports()
        {
            bool blnFatalError = false;
            int intCounter;
            int intNumberOfRecords;
            int intEmployeeID = -1;
            decimal decTotalHours = 0;
            decimal decMultiplier = 1;
            decimal decTotalCost = 0;
            decimal decReportedHours;
            decimal decPayRate;
            decimal decDifference;
            int intProjectCounter;
            bool blnItemFound;
            int intProjectID;
            bool blnMonday;
            DateTime datTransactionDate;
            DateTime datStartDate;
            DateTime datEndDate;
            string strFileName;
            string strPath = "\\\\bjc\\shares\\Documents\\ProductivityReports\\";
            string strHeader = "Project Productivity Report";
            string strMessage;
            string strEmailAddress;
            double douTotalPrice;

            try
            {
                datEndDate = DateTime.Now;
                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 90);
                TheEmployeeProductivityDataSet.employeeproductivity.Rows.Clear();
                TheCompleteProjectProductivityDataSet.completeprojectproductivity.Rows.Clear();

                TheFindAllEmployeesProductionOverAWeekDataSet = TheEmployeeProjectAssignmentClass.FindAllEmployeeProductionOverAWeek(datStartDate, datEndDate);

                intNumberOfRecords = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek.Rows.Count - 1;
                blnMonday = false;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    decPayRate = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].PayRate;
                    decReportedHours = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].TotalHours;
                    datTransactionDate = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].TransactionDate;

                    if ((datTransactionDate.DayOfWeek == DayOfWeek.Monday) && (blnMonday == false))
                    {
                        decTotalHours = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].TotalHours;
                        intEmployeeID = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].EmployeeID;
                        blnMonday = true;
                        decMultiplier = 1;
                    }
                    else if (intEmployeeID != TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].EmployeeID)
                    {
                        decTotalHours = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].TotalHours;
                        intEmployeeID = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].EmployeeID;
                        decMultiplier = 1;
                    }
                    else if (intEmployeeID == TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].EmployeeID)
                    {
                        decTotalHours += TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].TotalHours;

                        if (datTransactionDate.DayOfWeek != DayOfWeek.Monday)
                        {
                            blnMonday = false;
                        }
                    }

                    if (decMultiplier == 1)
                    {
                        if (decTotalHours <= 40)
                        {
                            decTotalCost = decPayRate * decReportedHours;
                        }
                        if (decTotalHours > 40)
                        {
                            decDifference = decTotalHours - 40;
                            decMultiplier = Convert.ToDecimal(1.5);
                            decTotalCost = ((decReportedHours - decDifference) * decPayRate) + (decDifference * decPayRate * decMultiplier);
                        }
                    }
                    if (decMultiplier == Convert.ToDecimal(1.5))
                    {
                        decTotalCost = decReportedHours * decPayRate * decMultiplier;
                    }

                    EmployeeProductivityDataSet.employeeproductivityRow NewProductivityRow = TheEmployeeProductivityDataSet.employeeproductivity.NewemployeeproductivityRow();

                    NewProductivityRow.AssignedProjectID = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].AssignedProjectID;
                    NewProductivityRow.CalculatedHours = decTotalHours;
                    NewProductivityRow.EmployeeID = intEmployeeID;
                    NewProductivityRow.FirstName = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].FirstName;
                    NewProductivityRow.LastName = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].LastName;
                    NewProductivityRow.Multiplier = decMultiplier;
                    NewProductivityRow.PayRate = decPayRate;
                    NewProductivityRow.ProjectID = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].ProjectID;
                    NewProductivityRow.ProjectName = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].ProjectName;
                    NewProductivityRow.TotalCost = decTotalCost;
                    NewProductivityRow.TotalHours = decReportedHours;
                    NewProductivityRow.TransactionDate = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].TransactionDate;

                    TheEmployeeProductivityDataSet.employeeproductivity.Rows.Add(NewProductivityRow);
                }

                TheFindAllDesignEmployeeProductivityOverAWeekDataSet = TheDesignProductivityClass.FindAllDesignEmployeeProductivityOverAWeek(datStartDate, datEndDate);

                intNumberOfRecords = TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    decPayRate = TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].PayRate;
                    decReportedHours = TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].TotalHours;
                    datTransactionDate = TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].TransactionDate;

                    if ((datTransactionDate.DayOfWeek == DayOfWeek.Monday) && (blnMonday == false))
                    {
                        decTotalHours = TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].TotalHours;
                        intEmployeeID = TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].EmployeeID;
                        blnMonday = true;
                        decMultiplier = 1;
                    }
                    else if (intEmployeeID != TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].EmployeeID)
                    {
                        decTotalHours = TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].TotalHours;
                        intEmployeeID = TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].EmployeeID;
                        decMultiplier = 1;
                    }
                    else if (intEmployeeID == TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].EmployeeID)
                    {
                        decTotalHours += TheFindAllDesignEmployeeProductivityOverAWeekDataSet.FindAllDesignEmployeeProductivityOverAWeek[intCounter].TotalHours;

                        if (datTransactionDate.DayOfWeek != DayOfWeek.Monday)
                        {
                            blnMonday = false;
                        }
                    }

                    if (decMultiplier == 1)
                    {
                        if (decTotalHours <= 40)
                        {
                            decTotalCost = decPayRate * decReportedHours;
                        }
                        if (decTotalHours > 40)
                        {
                            decDifference = decTotalHours - 40;
                            decMultiplier = Convert.ToDecimal(1.5);
                            decTotalCost = ((decReportedHours - decDifference) * decPayRate) + (decDifference * decPayRate * decMultiplier);
                        }
                    }
                    if (decMultiplier == Convert.ToDecimal(1.5))
                    {
                        decTotalCost = decReportedHours * decPayRate * decMultiplier;
                    }

                    EmployeeProductivityDataSet.employeeproductivityRow NewProductivityRow = TheEmployeeProductivityDataSet.employeeproductivity.NewemployeeproductivityRow();

                    NewProductivityRow.AssignedProjectID = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].AssignedProjectID;
                    NewProductivityRow.CalculatedHours = decTotalHours;
                    NewProductivityRow.EmployeeID = intEmployeeID;
                    NewProductivityRow.FirstName = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].FirstName;
                    NewProductivityRow.LastName = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].LastName;
                    NewProductivityRow.Multiplier = decMultiplier;
                    NewProductivityRow.PayRate = decPayRate;
                    NewProductivityRow.ProjectID = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].ProjectID;
                    NewProductivityRow.ProjectName = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].ProjectName;
                    NewProductivityRow.TotalCost = decTotalCost;
                    NewProductivityRow.TotalHours = decReportedHours;
                    NewProductivityRow.TransactionDate = TheFindAllEmployeesProductionOverAWeekDataSet.FindAllEmployeeProductionOverAWeek[intCounter].TransactionDate;

                    TheEmployeeProductivityDataSet.employeeproductivity.Rows.Add(NewProductivityRow);
                }

                intNumberOfRecords = TheEmployeeProductivityDataSet.employeeproductivity.Rows.Count - 1;
                gintProjectCounter = 0;
                gintProjectNumberOfRecords = 0;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    blnItemFound = false;
                    intProjectID = TheEmployeeProductivityDataSet.employeeproductivity[intCounter].ProjectID;
                    decTotalHours = TheEmployeeProductivityDataSet.employeeproductivity[intCounter].TotalHours;
                    decTotalCost = TheEmployeeProductivityDataSet.employeeproductivity[intCounter].TotalCost;

                    if (gintProjectCounter > 0)
                    {
                        for (intProjectCounter = 0; intProjectCounter <= gintProjectNumberOfRecords; intProjectCounter++)
                        {
                            if (intProjectID == TheCompleteProjectProductivityDataSet.completeprojectproductivity[intProjectCounter].ProjectID)
                            {
                                TheCompleteProjectProductivityDataSet.completeprojectproductivity[intProjectCounter].TotalHours += decTotalHours;
                                TheCompleteProjectProductivityDataSet.completeprojectproductivity[intProjectCounter].TotalCosts += decTotalCost;
                                blnItemFound = true;
                            }
                        }
                    }

                    if (blnItemFound == false)
                    {
                        CompleteProjectProductivityDataSet.completeprojectproductivityRow NewProjectRow = TheCompleteProjectProductivityDataSet.completeprojectproductivity.NewcompleteprojectproductivityRow();

                        NewProjectRow.AssignedProjectID = TheEmployeeProductivityDataSet.employeeproductivity[intCounter].AssignedProjectID;
                        NewProjectRow.ProjectID = intProjectID;
                        NewProjectRow.ProjectName = TheEmployeeProductivityDataSet.employeeproductivity[intCounter].ProjectName;
                        NewProjectRow.TotalCosts = decTotalCost;
                        NewProjectRow.TotalHours = decTotalHours;

                        TheCompleteProjectProductivityDataSet.completeprojectproductivity.Rows.Add(NewProjectRow);
                        gintProjectNumberOfRecords = gintProjectCounter;
                        gintProjectCounter++;
                    }
                }

                strMessage = "<h1>Project Productivity Report</h1>";
                strMessage += "<h1>An Excel Copy of the Report Can Be Found At " + strPath + "</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Assigned PProject ID</b></td>";
                strMessage += "<td><b>Project Name</b></td>";
                strMessage += "<td><b>Total Hours</b></td>";
                strMessage += "<td><b>Labor Costs</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";


                intNumberOfRecords = TheCompleteProjectProductivityDataSet.completeprojectproductivity.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strMessage += "<tr>";
                    strMessage += "<td><b>" + TheCompleteProjectProductivityDataSet.completeprojectproductivity[intCounter].AssignedProjectID + "</b></td>";
                    strMessage += "<td><b>" + TheCompleteProjectProductivityDataSet.completeprojectproductivity[intCounter].ProjectName + "</b></td>";
                    strMessage += "<td><b>" + Convert.ToString(TheCompleteProjectProductivityDataSet.completeprojectproductivity[intCounter].TotalHours) + "</b></td>";
                    strMessage += "<td><b>" + Convert.ToString(TheCompleteProjectProductivityDataSet.completeprojectproductivity[intCounter].TotalCosts) +  "</b></td>";
                    strMessage += "</tr>";
                }

                strMessage += "<table>";

                strEmailAddress = "ERPProjectReports@bluejaycommunications.com";

                blnFatalError = (TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage));

                if (blnFatalError == true)
                    throw new Exception();

                intNumberOfRecords = TheFindProductivityManagersForEmailDataSet.FindProductivityManagersForEmail.Rows.Count - 1;

                strFileName = "Productivity Report For " + Convert.ToString(datEndDate.Month) + "-" + Convert.ToString(datEndDate.Day) + "-" + Convert.ToString(datEndDate.Year);

                blnFatalError = ExportToExcel(strFileName, strPath);

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
                intRowNumberOfRecords = TheCompleteProjectProductivityDataSet.completeprojectproductivity.Rows.Count;
                intColumnNumberOfRecords = TheCompleteProjectProductivityDataSet.completeprojectproductivity.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheCompleteProjectProductivityDataSet.completeprojectproductivity.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheCompleteProjectProductivityDataSet.completeprojectproductivity.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                workbook.SaveAs(strPath + strFileName + ".xlsx");

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
