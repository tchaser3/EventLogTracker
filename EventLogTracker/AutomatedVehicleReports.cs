/* Title:           Automated Vehicle Reports
 * Date:            4-25-18
 * Author:          Terry Holmes
 * 
 * Description:     This the autommated report for emailing */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DateSearchDLL;
using InspectionsDLL;
using NewEventLogDLL;
using VehicleMainDLL;
using DataValidationDLL;
using NewEmployeeDLL;
using VehicleInYardDLL;
using VehiclesInShopDLL;
using VehicleAssignmentDLL;
using WeeklyBulkToolInspectionDLL;
using WeeklyInspectionsDLL;
using WeeklyVehicleCleanliness;
using EmployeeProjectAssignmentDLL;
using GEOFenceDLL;

namespace EventLogTracker
{
    class AutomatedVehicleReports
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        DateSearchClass TheDataSearchClass = new DateSearchClass();
        InspectionsClass TheInspectionClass = new InspectionsClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        VehicleMainClass TheVehicleMainClass = new VehicleMainClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        SendEmailClass TheSendEmailClass = new SendEmailClass();
        VehicleInYardClass TheVehicleInYardClass = new VehicleInYardClass();
        VehicleAssignmentClass TheVehicleAssignmentClass = new VehicleAssignmentClass();
        VehiclesInShopClass TheVehicleInShopClass = new VehiclesInShopClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        WeeklyInspectionClass TheWeeklyInspectionClass = new WeeklyInspectionClass();
        WeeklyVehicleCleanlinessClass TheWeeklyVehicleCleanlinessClass = new WeeklyVehicleCleanlinessClass();
        WeeklyBulkToolInspectionClass TheWeeklyBulkToolInspectionClass = new WeeklyBulkToolInspectionClass();
        EmployeeProjectAssignmentClass TheEmployeeProjectAssignmentClass = new EmployeeProjectAssignmentClass();
        GEOFenceClass TheGEOFenceClass = new GEOFenceClass();

        //setting up the data
        FindActiveVehicleMainByVehicleNumberDataSet TheFindActiveVehicleMainByVehicleNumberDataSet = new FindActiveVehicleMainByVehicleNumberDataSet();
        //FindDailyVehicleInspectionByDateRangeDataSet TheFindDailyVehicleInspectionByDateRangeDataSet = new FindDailyVehicleInspectionByDateRangeDataSet();
        FindVehicleInspectionProblemsByInspectionIDDataSet TheFindVehicleInspectionProblemByInsepctionIDDataSet = new FindVehicleInspectionProblemsByInspectionIDDataSet();
        FindDailyVehicleInspectionByVehicleIDAndDateRangeDataSet TheFindDailyVehicleInspectionByVehicleIDAndDateRangeDataSet = new FindDailyVehicleInspectionByVehicleIDAndDateRangeDataSet();
        FindDailyVehicleInspectionsByEmployeeIDAndDateRangeDataSet TheFindDailyVehicleInspectionByEmployeeIDAndDateRangeDataSet = new FindDailyVehicleInspectionsByEmployeeIDAndDateRangeDataSet();
        DailyVehicleInspectionReportDataSet TheDailyVehicleInspectionReportDataSet = new DailyVehicleInspectionReportDataSet();
        ComboEmployeeDataSet TheComboEmployeeDataSet = new ComboEmployeeDataSet();
        FindActiveVehicleMainDataSet TheFindActiveVehicleMainDataSet = new FindActiveVehicleMainDataSet();
        VehicleExceptionDataSet TheVehicleExceptionDataSet = new VehicleExceptionDataSet();
        FindVehiclesInYardByVehicleIDAndDateRangeDataSet TheFindVehicleInYardByVehicleIDAndDateRangeDataSet = new FindVehiclesInYardByVehicleIDAndDateRangeDataSet();
        FindVehicleMainInShopByVehicleIDDataSet TheFindVehicleMainInShopByVehicleIDDataSet = new FindVehicleMainInShopByVehicleIDDataSet();
        FindCurrentAssignedVehicleMainByVehicleIDDataSet TheFindcurrentAssignedVehicleMainByVehicleIDDataSet = new FindCurrentAssignedVehicleMainByVehicleIDDataSet();
        FindWeeklyVehicleMainInspectionByDateRangeDataSet TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet = new FindWeeklyVehicleMainInspectionByDateRangeDataSet();
        FindWeeklyVehicleMainInspectionProblemByInspectionIDDataSet TheFindWeeklyVehicleMainInspectionProblemsByInspectionIDDataSet = new FindWeeklyVehicleMainInspectionProblemByInspectionIDDataSet();
        FindWeeklyVehicleCleanlinessByInspectionIDDataSet TheFindWeeklyVehicleCleanlinessByInspectionIDDataSet = new FindWeeklyVehicleCleanlinessByInspectionIDDataSet();
        FindWeeklyBulkToolInspectionByInspectionIDDataSet TheFindWeeklyBulkToolInspectionByInspectionIDDataSet = new FindWeeklyBulkToolInspectionByInspectionIDDataSet();
        FindVehiclesInYardSummaryDataSet TheFindVehiclesInYardSummaryDataSet = new FindVehiclesInYardSummaryDataSet();
        DailyVehicleInspectionSummaryReportDataSet TheDailyVehicleInspectionSummaryReportDataSet = new DailyVehicleInspectionSummaryReportDataSet();
        FindEmployeeByEmployeeIDDataSet TheFindEmployeeByEmployeeIDDataSet = new FindEmployeeByEmployeeIDDataSet();
        FindWeeklyVehicleInspectionNoProblemDataSet TheFindWeeklyVehicleInspectionNoProblemDataSet = new FindWeeklyVehicleInspectionNoProblemDataSet();
        FindVehiclesInYardByDateRangeDataSet TheFindVehiclesInYardByDateRangeDataSet = new FindVehiclesInYardByDateRangeDataSet();
        FindProductivityNotCorrectDataSet TheFindProductivityNotCorrectDataSet = new FindProductivityNotCorrectDataSet();
        FindProductivityManagersForEmailDataSet TheFindProductivityManagersForEmailDataSet = new FindProductivityManagersForEmailDataSet();
        FindGEOFenceNoDriverAssignedDataSet TheFindGEOFenceNoDriverAssignedDataSet = new FindGEOFenceNoDriverAssignedDataSet();

        //settup created data set
        AuditReportDataSet TheAuditReportDataSet = new AuditReportDataSet();

        int gintEmployeeID;
        int gintVehicleID;
        DateTime gdatStartDate;
        DateTime gdatEndDate;
        string gstrReportType;

        public void RunAutomatedReports(DateTime datStartDate)
        {
            bool blnFatalError = false;

            try
            {
                blnFatalError = RunDailyVehicleInspectionReport(datStartDate);

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = EmailDailyVehicleInspectionReport();

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = RunGEOFenceReport();

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = RunVehicleExceptionReport(datStartDate);

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = EmailVehicleExceptionReport();

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = RunVehiclesInYardReport(datStartDate);

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = RunProductivityNotCorrect();

                if (blnFatalError == true)
                    throw new Exception();

            }
            catch(Exception Ex)
            {
                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private bool RunGEOFenceReport()
        {
            bool blnFatalError = false;
            DateTime datStartDate;
            DateTime datEndDate;
            int intCounter;
            int intNumberOfRecords;
            string strMessage;
            string strHeader = "Automile No Driver Assigned";
            int intManagerID;
            string strManagerName;
            string strEmailAddress;


            try
            {
                datStartDate = DateTime.Now;
                datStartDate = TheDataSearchClass.RemoveTime(datStartDate);
                datEndDate = TheDataSearchClass.AddingDays(datStartDate, 1);

                TheFindGEOFenceNoDriverAssignedDataSet = TheGEOFenceClass.FindGEOFenceNoDriverAssigned(datStartDate, datEndDate);

                intNumberOfRecords = TheFindGEOFenceNoDriverAssignedDataSet.FindGEOFenceNoDriverAssigned.Rows.Count - 1;

                strMessage = "<h1>Automile No Driver Assigned</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Transaction ID</b></td>";
                strMessage += "<td><b>Date</b></td>";
                strMessage += "<td><b>Vehicle Number</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Assigned Office</b></td>";
                strMessage += "<td><b>Driver</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                if (intNumberOfRecords > -1)
                {
                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        //adding items to message
                        strMessage += "<tr>";
                        strMessage += "<td>" + Convert.ToString(TheFindGEOFenceNoDriverAssignedDataSet.FindGEOFenceNoDriverAssigned[intCounter].TransactionID) + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheFindGEOFenceNoDriverAssignedDataSet.FindGEOFenceNoDriverAssigned[intCounter].EventDate) + "</td>";
                        strMessage += "<td>" + TheFindGEOFenceNoDriverAssignedDataSet.FindGEOFenceNoDriverAssigned[intCounter].VehicleNumber+ "</td>";
                        strMessage += "<td>" + TheFindGEOFenceNoDriverAssignedDataSet.FindGEOFenceNoDriverAssigned[intCounter].FirstName + "</td>";
                        strMessage += "<td>" + TheFindGEOFenceNoDriverAssignedDataSet.FindGEOFenceNoDriverAssigned[intCounter].LastName + "</td>";
                        strMessage += "<td>" + TheFindGEOFenceNoDriverAssignedDataSet.FindGEOFenceNoDriverAssigned[intCounter].AssignedOffice + "</td>";
                        strMessage += "<td>" + TheFindGEOFenceNoDriverAssignedDataSet.FindGEOFenceNoDriverAssigned[intCounter].Driver + "</td>";
                        strMessage += "</tr>";
                    }
                }

                TheFindProductivityManagersForEmailDataSet = TheEmployeeProjectAssignmentClass.FindProductivityManagersForEmail();

                intNumberOfRecords = TheFindProductivityManagersForEmailDataSet.FindProductivityManagersForEmail.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)

                {
                    strEmailAddress = TheFindProductivityManagersForEmailDataSet.FindProductivityManagersForEmail[intCounter].EmailAddress;

                    blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                    if (blnFatalError == true)
                        throw new Exception();
                }

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automated Vehicle Reprots // Run GEO Fence Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;

            }


            return blnFatalError;
        }
        private bool RunProductivityNotCorrect()
        {
            bool blnFatalError = false;
            DateTime datStartDate;
            DateTime datEndDate;
            int intCounter;
            int intNumberOfRecords;
            string strMessage;
            string strHeader = "Productivity Sheets Not Correct"; 
            int intManagerID;
            string strManagerName;
            string strEmailAddress;


            try
            {
                datEndDate = DateTime.Now;
                datEndDate = TheDataSearchClass.RemoveTime(datEndDate);
                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 1);

                TheFindProductivityNotCorrectDataSet = TheEmployeeProjectAssignmentClass.FindProductivityNotCorrect(datStartDate, datEndDate);
                intNumberOfRecords = TheFindProductivityNotCorrectDataSet.FindProductivityNotCorrect.Rows.Count - 1;

                strMessage = "<h1>Productivity Sheets Not Correct</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Transaction ID</b></td>";
                strMessage += "<td><b>Date</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Assigned Office</b></td>";
                strMessage += "<td><b>Manager</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                if(intNumberOfRecords > -1)
                {
                    for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        intManagerID = TheFindProductivityNotCorrectDataSet.FindProductivityNotCorrect[intCounter].ManagerID;

                        TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);
                        strManagerName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName + " ";
                        strManagerName += TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;


                        //adding items to message
                        strMessage += "<tr>";
                        strMessage += "<td>" + Convert.ToString(TheFindProductivityNotCorrectDataSet.FindProductivityNotCorrect[intCounter].TransactionID) + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheFindProductivityNotCorrectDataSet.FindProductivityNotCorrect[intCounter].TransactionDate) + "</td>";
                        strMessage += "<td>" + TheFindProductivityNotCorrectDataSet.FindProductivityNotCorrect[intCounter].FirstName + "</td>";
                        strMessage += "<td>" + TheFindProductivityNotCorrectDataSet.FindProductivityNotCorrect[intCounter].LastName + "</td>";
                        strMessage += "<td>" + TheFindProductivityNotCorrectDataSet.FindProductivityNotCorrect[intCounter].HomeOffice + "</td>";
                        strMessage += "<td>" + strManagerName + "</td>";
                        strMessage += "</tr>";
                    }
                }

                TheFindProductivityManagersForEmailDataSet = TheEmployeeProjectAssignmentClass.FindProductivityManagersForEmail();

                intNumberOfRecords = TheFindProductivityManagersForEmailDataSet.FindProductivityManagersForEmail.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strEmailAddress = TheFindProductivityManagersForEmailDataSet.FindProductivityManagersForEmail[intCounter].EmailAddress;

                    blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                    if (blnFatalError == true)
                        throw new Exception();
                }

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automated Vehicle Reprots // Run Productivity Not Correct " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;

            }

            return blnFatalError;
        }


        private bool RunVehiclesInYardReport(DateTime datStartDate)
        {
            bool blnFatalError = false;
            int intCounter;
            int intNumberOfRecords;
            string strMessage = "";
            string strHeader = "Vehicles in Yard for Today";
            DateTime datEndDate = DateTime.Now;
            int intManagerID = 0;
            string strManagerName = "";

            try
            {
                datStartDate = TheDateSearchClass.RemoveTime(datStartDate);
                datEndDate = TheDateSearchClass.AddingDays(datStartDate, 1);
                TheFindVehiclesInYardByDateRangeDataSet = TheVehicleInYardClass.FindVehiclesInYardByDateRange(datStartDate, datEndDate);
                intNumberOfRecords = TheFindVehiclesInYardByDateRangeDataSet.FindVehiclesInYardByDateRange.Rows.Count - 1;

                strMessage = "<h1>Vehicles in Yard</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Vehicle Number</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Assigned Office</b></td>";
                strMessage += "<td><b>Manager</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                if(intNumberOfRecords > -1)
                {
                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        intManagerID = TheFindVehiclesInYardByDateRangeDataSet.FindVehiclesInYardByDateRange[intCounter].ManagerID;

                        TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);
                        strManagerName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName + " ";
                        strManagerName += TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;

                        if(TheFindVehiclesInYardByDateRangeDataSet.FindVehiclesInYardByDateRange[intCounter].LastName == "WAREHOUSE")
                        {
                            strManagerName = "FLEET MANAGER";
                        }

                        //adding items to message
                        strMessage += "<tr>";
                        strMessage += "<td>" + TheFindVehiclesInYardByDateRangeDataSet.FindVehiclesInYardByDateRange[intCounter].VehicleNumber + "</td>";
                        strMessage += "<td>" + TheFindVehiclesInYardByDateRangeDataSet.FindVehiclesInYardByDateRange[intCounter].FirstName + "</td>";
                        strMessage += "<td>" + TheFindVehiclesInYardByDateRangeDataSet.FindVehiclesInYardByDateRange[intCounter].LastName + "</td>";
                        strMessage += "<td>" + TheFindVehiclesInYardByDateRangeDataSet.FindVehiclesInYardByDateRange[intCounter].AssignedOffice + "</td>";
                        strMessage += "<td>" + strManagerName + "</td>";
                        strMessage += "</tr>";

                    }
                }

                TheSendEmailClass.VehicleReports(strHeader, strMessage);
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automated Reports // Run Vehicles In Yard Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }


            return blnFatalError;
        }
        private void RunNoProblemReport()
        {
            int intCounter;
            int intNumberOfRecords;
            string strMessage = "";
            string strHeader = "Weekly Inspection No Problem Vs Open Problem Report - Do Not Reply";
            DateTime datStartDate;
            DateTime datEndDate = DateTime.Now;
            bool blnFatalError = false;

            try
            {
                datEndDate = TheDataSearchClass.RemoveTime(datEndDate);
                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 7);
                TheFindWeeklyVehicleInspectionNoProblemDataSet = TheWeeklyInspectionClass.FindWeeklyVehicleInspectionNoProblem(datStartDate, datEndDate);
                intNumberOfRecords = TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem.Rows.Count - 1;

                strMessage = "<h1>Weekly Inspection No Problem Vs Open Problem Report</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Problem ID</b></td>";
                strMessage += "<td><b>Vehicle Number</b></td>";
                strMessage += "<td><b>Problem Date</b></td>";
                strMessage += "<td><b>Problem</b></td>";
                strMessage += "<td><b>Inspection Date</b></td>";
                strMessage += "<td><b>Inspector First Name</b></td>";
                strMessage += "<td><b>Inspector Last Name</b></td>";
                strMessage += "<td><b>Inspection Status</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strMessage += "<tr>";
                    strMessage += "<td>" + Convert.ToString(TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem[intCounter].ProblemID) + "</td>";
                    strMessage += "<td>" + TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem[intCounter].VehicleNumber + "</td>";
                    strMessage += "<td>" + Convert.ToString(TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem[intCounter].TransactionDate) + "</td>";
                    strMessage += "<td>" + TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem[intCounter].Problem + "</td>";
                    strMessage += "<td>" + Convert.ToString(TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem[intCounter].InspectionDate) + "</td>";
                    strMessage += "<td>" + TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem[intCounter].FirstName + "</td>";
                    strMessage += "<td>" + TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem[intCounter].LastName + "</td>";
                    strMessage += "<td>" + TheFindWeeklyVehicleInspectionNoProblemDataSet.FindWeeklyVehicleInspectionNoProblem[intCounter].InspectionStatus + "</td>";
                    strMessage += "</tr>";
                }

                strMessage += "<table>";

                blnFatalError = TheSendEmailClass.VehicleReports(strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automate Vehicle Reports // /Run No Problem Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private bool EmailVehicleExceptionReport()
        {
            bool blnFatalError = false;
            int intCounter;
            int intNumberOfRecords;
            string strMessage = "";
            string strHeader = "Vehicle Exception Report - Do Not Reply";

            try
            {
                intNumberOfRecords = TheVehicleExceptionDataSet.vehicleexception.Rows.Count - 1;

                strMessage = "<h1>Vehicle Exception Report<h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Vehicle Number</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Home Office</b></td>";
                strMessage += "<td><b>Manager</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strMessage += "<tr>";
                    strMessage += "<td>" + TheVehicleExceptionDataSet.vehicleexception[intCounter].VehicleNumber + "</td>";
                    strMessage += "<td>" + TheVehicleExceptionDataSet.vehicleexception[intCounter].FirstName + "</td>";
                    strMessage += "<td>" + TheVehicleExceptionDataSet.vehicleexception[intCounter].LastName + "</td>";
                    strMessage += "<td>" + TheVehicleExceptionDataSet.vehicleexception[intCounter].AssignedOffice + "</td>";
                    strMessage += "<td>" + TheVehicleExceptionDataSet.vehicleexception[intCounter].Manager + "</td>";
                    strMessage += "</tr>";
                }

                strMessage += "<table>";

                blnFatalError = TheSendEmailClass.VehicleReports(strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay ERP // Automate Vehicle Reports // Email Vehicle Exception Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
        private bool RunVehicleExceptionReport(DateTime datStartDate)
        {
            bool blnFatalError = false;
            int intCounter;
            int intNumberOfRecords;
            int intVehicleID;
            int intRecordsReturned;
            DateTime datEndDate;
            string strFirstNamed = "";
            string strLastName = "";
            string strManager = "";
            int intManagerID = 0;

            try
            {
                TheVehicleExceptionDataSet.vehicleexception.Rows.Clear();

                datEndDate = TheDateSearchClass.AddingDays(datStartDate, 1);

                TheFindActiveVehicleMainDataSet = TheVehicleMainClass.FindActiveVehicleMain();

                intNumberOfRecords = TheFindActiveVehicleMainDataSet.FindActiveVehicleMain.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intVehicleID = TheFindActiveVehicleMainDataSet.FindActiveVehicleMain[intCounter].VehicleID;

                    TheFindDailyVehicleInspectionByVehicleIDAndDateRangeDataSet = TheInspectionClass.FindDailyVehicleInspectionByVehicleIDAndDateRange(intVehicleID, datStartDate, datEndDate);

                    intRecordsReturned = TheFindDailyVehicleInspectionByVehicleIDAndDateRangeDataSet.FindDailyVehicleInspectionsByVehicleIDAndDateRange.Rows.Count;

                    if (intRecordsReturned == 0)
                    {
                        TheFindVehicleInYardByVehicleIDAndDateRangeDataSet = TheVehicleInYardClass.FindVehiclesInYardByVehicleIDAndDateRange(intVehicleID, datStartDate, datEndDate);

                        intRecordsReturned = TheFindVehicleInYardByVehicleIDAndDateRangeDataSet.FindVehiclesInYardByVehicleIDAndDateRange.Rows.Count;

                        if (intRecordsReturned == 0)
                        {
                            TheFindVehicleMainInShopByVehicleIDDataSet = TheVehicleInShopClass.FindVehicleMainInShopByVehicleID(intVehicleID);

                            intRecordsReturned = TheFindVehicleMainInShopByVehicleIDDataSet.FindVehicleMainInShopByVehicleID.Rows.Count;

                            if (intRecordsReturned == 0)
                            {
                                TheFindcurrentAssignedVehicleMainByVehicleIDDataSet = TheVehicleAssignmentClass.FindCurrentAssignedVehicleMainByVehicleID(intVehicleID);

                                intRecordsReturned = TheFindcurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID.Rows.Count;

                                if (intRecordsReturned == 0)
                                {
                                    strLastName = "NOT ASSIGNED";
                                    strFirstNamed = "NOT ASSIGNED";
                                    strManager = "FLEET MANAGER";
                                }
                                else
                                {
                                    strLastName = TheFindcurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID[0].LastName;
                                    strFirstNamed = TheFindcurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID[0].FirstName;

                                    intManagerID = TheFindcurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID[0].ManagerID;

                                    TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);

                                    strManager = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName + " " + TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;
                                }

                                VehicleExceptionDataSet.vehicleexceptionRow NewVehicleRow = TheVehicleExceptionDataSet.vehicleexception.NewvehicleexceptionRow();

                                NewVehicleRow.AssignedOffice = TheFindActiveVehicleMainDataSet.FindActiveVehicleMain[intCounter].AssignedOffice;
                                NewVehicleRow.FirstName = strFirstNamed;
                                NewVehicleRow.LastName = strLastName;
                                NewVehicleRow.VehicleID = intVehicleID;
                                NewVehicleRow.VehicleNumber = TheFindActiveVehicleMainDataSet.FindActiveVehicleMain[intCounter].VehicleNumber;

                                if(TheFindActiveVehicleMainDataSet.FindActiveVehicleMain[intCounter].LastName == "WAREHOUSE")
                                {
                                    strManager = "FLEET MANAGER";
                                }

                                NewVehicleRow.Manager = strManager;

                                TheVehicleExceptionDataSet.vehicleexception.Rows.Add(NewVehicleRow);
                            }
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automate Vehicle Reports // Run Vehicle Exception Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
        public bool RunWeeklyVehicleInspectionReport()
        {
            bool blnFatalError = false;
            DateTime datStartDate = DateTime.Now;
            DateTime datEndDate = DateTime.Now;
            int intRecordsReturned;
            int intCounter;
            int intNumberOfRecords;
            int intInspectionID;
            string strNotes;
            string strCleanlinessNotes;
            string strToolNotes;
            string strCones = "INCORRECT";
            string strSigns = "INCORRECT";
            string strFirstAidKits = "INCORRECT";
            string strFireExtinguishers = "INCORRECT";

            try
            {
                TheAuditReportDataSet.auditreport.Rows.Clear();
                datEndDate = TheDateSearchClass.RemoveTime(datEndDate);
                datEndDate = TheDateSearchClass.AddingDays(datEndDate, 1);
                datStartDate = TheDataSearchClass.SubtractingDays(datEndDate, 9);

                TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet = TheWeeklyInspectionClass.FindWeeklyVehicleMainInspectionByDateRange(datStartDate, datEndDate);

                intNumberOfRecords = TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet.FindWeeklyVehicleMainInspectionByDateRange.Rows.Count - 1;

                if (intNumberOfRecords > -1)
                {
                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        strNotes = " ";
                        intInspectionID = TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet.FindWeeklyVehicleMainInspectionByDateRange[intCounter].TransactionID;

                        TheFindWeeklyVehicleMainInspectionProblemsByInspectionIDDataSet = TheWeeklyInspectionClass.FindWeeklyVehicleMainInspectionProblemByInspectionID(intInspectionID);

                        TheFindWeeklyVehicleCleanlinessByInspectionIDDataSet = TheWeeklyVehicleCleanlinessClass.FindWeeklyVehicleCleanlinessbyInspectionID(intInspectionID);

                        intRecordsReturned = TheFindWeeklyVehicleCleanlinessByInspectionIDDataSet.FindWeeklyVehicleCleanlinessByInspectionID.Rows.Count;

                        if (intRecordsReturned > 0)
                        {
                            strCleanlinessNotes = TheFindWeeklyVehicleCleanlinessByInspectionIDDataSet.FindWeeklyVehicleCleanlinessByInspectionID[0].CleanlinessNotes;
                        }
                        else
                        {
                            strCleanlinessNotes = "NO NOTES ENTERED";
                        }

                        intRecordsReturned = TheFindWeeklyVehicleMainInspectionProblemsByInspectionIDDataSet.FindWeeklyVehicleMainInspectionProblemByInspectionID.Rows.Count;

                        if (intRecordsReturned > 0)
                        {
                            strNotes = TheFindWeeklyVehicleMainInspectionProblemsByInspectionIDDataSet.FindWeeklyVehicleMainInspectionProblemByInspectionID[0].VehicleProblem;
                        }

                        TheFindWeeklyBulkToolInspectionByInspectionIDDataSet = TheWeeklyBulkToolInspectionClass.FindWeeklyBulkToolInspectionByInspectionID(intInspectionID);

                        intRecordsReturned = TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID.Rows.Count;

                        if (intRecordsReturned < 1)
                        {
                            strToolNotes = "NO INSPECTION NOTES WERE FOUND";
                            strCones = "UNKNOWN";
                            strSigns = "UNKNOW";
                            strFirstAidKits = "UNKNOWN";
                            strFireExtinguishers = "UNKNOWN";
                        }
                        else
                        {
                            strCones = "INCORRECT";
                            strSigns = "INCORRECT";
                            strFirstAidKits = "INCORRECT";
                            strFireExtinguishers = "INCORRECT";

                            if (TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].IsConesCorrectNull() == true)
                            {
                                strCones = "UNKNOWN";
                            }
                            else
                            {
                                if (TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].ConesCorrect == true)
                                    strCones = "CORRECT";
                            }

                            if (TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].IsSignsCorrectNull() == true)
                            {
                                strSigns = "UNKNOWN";
                            }
                            else
                            {
                                if(TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].SignsCorrect == true)
                                     strSigns = "CORRECT";
                            }

                            if (TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].IsFireExtingisherNull() == true)
                            {
                                strFireExtinguishers = "UNKNOWN";
                            }
                            else
                            {
                                if(TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].FireExtingisher == true)
                                    strFireExtinguishers = "CORRECT";
                            }

                            if (TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].IsFirstAidCorrectNull() == true)
                            {
                                strFirstAidKits = "UNKNOWN";
                            }
                            else
                            {
                                if(TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].FirstAidCorrect == true)
                                    strFirstAidKits = "CORRECT";
                            }

                            strToolNotes = TheFindWeeklyBulkToolInspectionByInspectionIDDataSet.FindWeeklyBulkToolInspectionByInspectionID[0].InspectionNotes;
                        }

                        AuditReportDataSet.auditreportRow NewInspectionRow = TheAuditReportDataSet.auditreport.NewauditreportRow();

                        NewInspectionRow.VehicleNumber = TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet.FindWeeklyVehicleMainInspectionByDateRange[intCounter].VehicleNumber;
                        NewInspectionRow.Date = TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet.FindWeeklyVehicleMainInspectionByDateRange[intCounter].InspectionDate;
                        NewInspectionRow.Findings = TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet.FindWeeklyVehicleMainInspectionByDateRange[intCounter].InspectionStatus;
                        NewInspectionRow.FirstName = TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet.FindWeeklyVehicleMainInspectionByDateRange[intCounter].FirstName;
                        NewInspectionRow.LastName = TheFindWeeklyVehicleMainInspectionbyDateRangeDataSet.FindWeeklyVehicleMainInspectionByDateRange[intCounter].LastName;
                        NewInspectionRow.Notes = strNotes;
                        NewInspectionRow.CleanlinessNotes = strCleanlinessNotes;
                        NewInspectionRow.ToolNotes = strToolNotes;
                        NewInspectionRow.ConesCorrect = strCones;
                        NewInspectionRow.SignsCorrect = strSigns;
                        NewInspectionRow.FirstAidCorrect = strFirstAidKits;
                        NewInspectionRow.FireExtinguisherCorrect = strFireExtinguishers;

                        TheAuditReportDataSet.auditreport.Rows.Add(NewInspectionRow);
                    }
                }

                EmailWeeklyReport();
                RunNoProblemReport();

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automated Vehicle Reports // Run Weekly Vehicle Inspection Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }


            return blnFatalError;
        }
        private void EmailWeeklyReport()
        {
            int intCounter;
            int intNumberOfRecords;
            string strMessage = "";
            string strHeader = "Manager/Weekly Vehicle Inspection Report - Do Not Reply";
            bool blnFatalError = false;

            try
            {
                intNumberOfRecords = TheAuditReportDataSet.auditreport.Rows.Count - 1;

                strMessage = "<h1>Manager/Weekly Vehicle Inspection Report<h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Vehicle Number</b></td>";
                strMessage += "<td><b>Date</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Findings</b></td>";
                strMessage += "<td><b>Notes</b></td>";
                strMessage += "<td><b>Cleanliness Notes</b></td>";
                strMessage += "<td><b>Cones Correct</b></td>";
                strMessage += "<td><b>Signs Correct</b></td>";
                strMessage += "<td><b>First Correct</b></td>";
                strMessage += "<td><b>Fire Correct</b></td>";
                strMessage += "<td><b>Tool Notes</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strMessage += "<tr>";
                    strMessage += "<td>" + TheAuditReportDataSet.auditreport[intCounter].VehicleNumber + "</td>";
                    strMessage += "<td>" + Convert.ToString(TheAuditReportDataSet.auditreport[intCounter].Date) + "</td>";
                    strMessage += "<td>" + TheAuditReportDataSet.auditreport[intCounter].FirstName + "</td>";
                    strMessage += "<td>" + TheAuditReportDataSet.auditreport[intCounter].LastName + "</td>";
                    strMessage += "<td>" + TheAuditReportDataSet.auditreport[intCounter].Findings + "</td>";
                    strMessage += "<td>" + TheAuditReportDataSet.auditreport[intCounter].Notes + "</td>";
                    strMessage += "<td>" + TheAuditReportDataSet.auditreport[intCounter].CleanlinessNotes + "</td>";
                    strMessage += "<td>" + Convert.ToString(TheAuditReportDataSet.auditreport[intCounter].ConesCorrect) + "</td>";
                    strMessage += "<td>" + Convert.ToString(TheAuditReportDataSet.auditreport[intCounter].SignsCorrect) + "</td>";
                    strMessage += "<td>" + Convert.ToString(TheAuditReportDataSet.auditreport[intCounter].FirstAidCorrect) + "</td>";
                    strMessage += "<td>" + Convert.ToString(TheAuditReportDataSet.auditreport[intCounter].FireExtinguisherCorrect) + "</td>";
                    strMessage += "<td>" + TheAuditReportDataSet.auditreport[intCounter].ToolNotes + "</td>";
                    strMessage += "</tr>";
                }

                strMessage += "<table>";

                blnFatalError = TheSendEmailClass.VehicleReports(strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay ERP // Automate Vehicle Reports // Email Vehicle Exception Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private bool RunDailyVehicleInspectionReport(DateTime datStartDate)
        {
            //setting local varaibles
             DateTime datLimitDate;
            int intCounter;
            int intNumberOfRecords;
            bool blnFatalError = false;
            int intManagerID;

            try
            {
                TheDailyVehicleInspectionReportDataSet.dailyinspection.Rows.Clear();
                datLimitDate = TheDataSearchClass.AddingDays(datStartDate, 1);

                TheDailyVehicleInspectionSummaryReportDataSet = TheInspectionClass.DailyVehicleInspectionSummaryReport(datStartDate, datLimitDate);

                intNumberOfRecords = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport.Rows.Count - 1;

                if (intNumberOfRecords > -1)
                {
                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        intManagerID = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport[intCounter].ManagerID;

                        TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);

                        DailyVehicleInspectionReportDataSet.dailyinspectionRow NewInspectionRow = TheDailyVehicleInspectionReportDataSet.dailyinspection.NewdailyinspectionRow();

                        NewInspectionRow.FirstName = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport[intCounter].FirstName;
                        NewInspectionRow.HomeOffice = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport[intCounter].HomeOffice;
                        NewInspectionRow.InspectionDate = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport[intCounter].InspectionDate;
                        NewInspectionRow.InspectionStatus = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport[intCounter].InspectionStatus;
                        NewInspectionRow.LastName = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport[intCounter].LastName;
                        NewInspectionRow.ManagerFirstName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName;
                        NewInspectionRow.ManagerLastName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;
                        NewInspectionRow.VehicleNumber = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport[intCounter].VehicleNumber;
                        NewInspectionRow.VehicleProblem = TheDailyVehicleInspectionSummaryReportDataSet.DailyVehicleInspectionSummaryReport[intCounter].VehicleProblem;

                        TheDailyVehicleInspectionReportDataSet.dailyinspection.Rows.Add(NewInspectionRow);
                    }
                }

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay ERP // Automated Vehicle Reports // Run Daily Vehicle Inspection Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
        private bool EmailDailyVehicleInspectionReport()
        {
            bool blnFatalError = false;
            int intCounter;
            int intNumberOfRecords;
            string strMessage = "<h1>Daily Vehicle Inspection Notes</h1>";
            string strHeader = "Daily Vehicle Inspection Report - Do Not Reply";

            try
            {
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Vehicle Number</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Home Office</b></td>";
                strMessage += "<td><b>Manager First Name</b></td>";
                strMessage += "<td><b>Manager Last Name</b></td>";
                strMessage += "<td><b>Inspection Status</b></td>";
                strMessage += "<td><b>Vehicle Problem</b></td>";                
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                //TodaysInspections();

                intNumberOfRecords = TheDailyVehicleInspectionReportDataSet.dailyinspection.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strMessage += "<tr>";
                    strMessage += "<td>" + TheDailyVehicleInspectionReportDataSet.dailyinspection[intCounter].VehicleNumber + "</td>";
                    strMessage += "<td>" + TheDailyVehicleInspectionReportDataSet.dailyinspection[intCounter].FirstName + "</td>";
                    strMessage += "<td>" + TheDailyVehicleInspectionReportDataSet.dailyinspection[intCounter].LastName + "</td>";
                    strMessage += "<td>" + TheDailyVehicleInspectionReportDataSet.dailyinspection[intCounter].HomeOffice + "</td>";
                    strMessage += "<td>" + TheDailyVehicleInspectionReportDataSet.dailyinspection[intCounter].ManagerFirstName + "</td>";
                    strMessage += "<td>" + TheDailyVehicleInspectionReportDataSet.dailyinspection[intCounter].ManagerLastName + "</td>";
                    strMessage += "<td>" + TheDailyVehicleInspectionReportDataSet.dailyinspection[intCounter].InspectionStatus + "</td>";
                    strMessage += "<td>" + TheDailyVehicleInspectionReportDataSet.dailyinspection[intCounter].VehicleProblem + "</td>";
                    strMessage += "</tr>";

                }

                strMessage += "</table>";

                blnFatalError =  TheSendEmailClass.VehicleReports(strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automated Vehicle Reports  // Email Daily Vehicle Inspection Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
        public void RunWeeklyVehiclesInYardReport()
        {
            //setting local variables
            int intCounter;
            int intNumberOfRecords;
            DateTime datStartDate = DateTime.Now;
            DateTime datEndDate = DateTime.Now;
            string strMessage = "<h1>Vehicles In The Yard Over the Last Week</h1>";
            string strHeader = "Vehicles In Yard Weekly Report - Do Not Reply";
            bool blnFatalError = false;

            try
            {
                datEndDate = TheDateSearchClass.RemoveTime(datEndDate);
                datEndDate = TheDateSearchClass.AddingDays(datEndDate, 1);
                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 7);

                TheFindVehiclesInYardSummaryDataSet = TheVehicleInYardClass.FindVehiclesInYardSummary(datStartDate, datEndDate);

                intNumberOfRecords = TheFindVehiclesInYardSummaryDataSet.FindVehiclesInYardSummary.Rows.Count - 1;

                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Vehicle Number</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Home Office</b></td>";
                strMessage += "<td><b>Times In Yard</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                //TodaysInspections();

                intNumberOfRecords = TheFindVehiclesInYardSummaryDataSet.FindVehiclesInYardSummary.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strMessage += "<tr>";
                    strMessage += "<td>" + TheFindVehiclesInYardSummaryDataSet.FindVehiclesInYardSummary[intCounter].VehicleNumber + "</td>";
                    strMessage += "<td>" + TheFindVehiclesInYardSummaryDataSet.FindVehiclesInYardSummary[intCounter].FirstName + "</td>";
                    strMessage += "<td>" + TheFindVehiclesInYardSummaryDataSet.FindVehiclesInYardSummary[intCounter].LastName + "</td>";
                    strMessage += "<td>" + TheFindVehiclesInYardSummaryDataSet.FindVehiclesInYardSummary[intCounter].HomeOffice + "</td>";
                    strMessage += "<td>" + Convert.ToString(TheFindVehiclesInYardSummaryDataSet.FindVehiclesInYardSummary[intCounter].TimesInYard) + "</td>";
                    strMessage += "</tr>";

                }

                strMessage += "</table>";

                blnFatalError = TheSendEmailClass.VehicleReports(strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Automated Vehicle Reports // Run Weekly Vehicles In Yard " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
