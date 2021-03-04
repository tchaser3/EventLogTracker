/* Title:           Event Log Tracker
 * Date:            3-23-17
 * Author:          Terry Holmes
 * 
 * Description:     This will show any exceptions thrown */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Net.Mail;
using System.Configuration;
using System.Net;
using DateSearchDLL;
using NewEventLogDLL;
using System.IO;
using NewVehicleDLL;
using VehicleProblemsDLL;
using VehiclesInShopDLL;
using VehicleStatusDLL;
using VehicleExceptionEmailDLL;
using NewEmployeeDLL;
using EmployeeLaborRateDLL;
using VehicleInYardDLL;
using VehicleAssignmentDLL;
using VehicleMainDLL;
using EmployeeProjectAssignmentDLL;
using EmployeeDateEntryDLL;
using ProjectMatrixDLL;
using System.Data;
using ProductionProjectDLL;
using GEOFenceDLL;

namespace EventLogTracker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        StreamWriter TheWriter;
        VehicleClass TheVehicleClass = new VehicleClass();
        VehicleProblemClass TheVehicleProblemClass = new VehicleProblemClass();
        VehiclesInShopClass TheVehicleInShopClass = new VehiclesInShopClass();
        VehicleStatusClass TheVehicleStatusClass = new VehicleStatusClass();
        UpdatingWorkTaskStatsClass TheUpdatingWorkTaskStatsClass = new UpdatingWorkTaskStatsClass();
        VehicleExceptionEmailClass TheVehicleExceptionEmailClass = new VehicleExceptionEmailClass();
        AutomatedVehicleReports TheAutomatedVehicleReportsClass = new AutomatedVehicleReports();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        EmployeeLaborRateClass TheEmployeeLaborRateClass = new EmployeeLaborRateClass();
        SendEmailClass TheSendEmailClass = new SendEmailClass();
        VehicleInYardClass TheVehicleInYardClass = new VehicleInYardClass();
        VehicleAssignmentClass TheVehicleAssignmentClass = new VehicleAssignmentClass();
        VehicleMainClass TheVehicleMainClass = new VehicleMainClass();
        EmployeeProjectAssignmentClass TheEmployeeProjectAssignmentClass = new EmployeeProjectAssignmentClass();
        AutomatedProductionReportsClass TheAutomatedProductioinReportsClass = new AutomatedProductionReportsClass();
        EmployeeDateEntryClass TheEmployeeDataEntryClass = new EmployeeDateEntryClass();
        ProjectMatrixClass TheProjectMatrixClass = new ProjectMatrixClass();
        ProductionProjectClass TheProductionProjectClass = new ProductionProjectClass();
        GEOFenceClass TheGeoFenceClass = new GEOFenceClass();

        //setting up the time
        DispatcherTimer MyTimer = new DispatcherTimer();
        DispatcherTimer AccessTimer = new DispatcherTimer();

        //the data set
        EventLogDataSet TheEventLogDataSet = new EventLogDataSet();
        FindEventLogByDateRangeDataSet TheFindEventLogByDateRangeDataSet = new FindEventLogByDateRangeDataSet();
        FindActiveVehiclesDataSet TheFindActiveVehiclesDataSet = new FindActiveVehiclesDataSet();
        FindOpenVehicleProblemsByVehicleIDDataSet TheFindOpenVehicleProblemsByVehicleIDDataSet = new FindOpenVehicleProblemsByVehicleIDDataSet();
        FindVehicleStatusByVehicleIDDataSet TheFindVehicleStatusByVehicleIDDataSet = new FindVehicleStatusByVehicleIDDataSet();
        FindVehicleMainInShopByVehicleIDDataSet TheFindVehicleMainInShopByVehicleIDDataSet = new FindVehicleMainInShopByVehicleIDDataSet();
        VehicleExceptionEmailDataSet TheVehicleExceptionEmailDataSet = new VehicleExceptionEmailDataSet();
        FindActiveEmployeesDataSet TheFindActiveEmployeesDataSet = new FindActiveEmployeesDataSet();
        FindEmployeeLaborRateDataSet TheFindEmployeeLaborRateDataSet = new FindEmployeeLaborRateDataSet();
        WeeklyVehicleReportsDateDataSet TheWeeklyVehicleReportsDateDataSet = new WeeklyVehicleReportsDateDataSet();
        FindWarehouseByWarehouseNameDataSet TheFindWarehouseByWarehouseNameDataSet = new FindWarehouseByWarehouseNameDataSet();
        FindVehiclesInyardShowingVehicleIDDateRangeDataSet TheFindVehiclesInYardShowingVehicleIDDateRangeDataSet = new FindVehiclesInyardShowingVehicleIDDateRangeDataSet();
        FindCurrentAssignedVehicleMainByVehicleIDDataSet TheFindCurrentAssignedVehicleMainByVehicleIDDataSet = new FindCurrentAssignedVehicleMainByVehicleIDDataSet();
        FindProductivityNotCorrectDataSet TheFindProductivityNotCorrectClass = new FindProductivityNotCorrectDataSet();
        FindEmployeeDataEntryByDateRangeDataSet TheFindEmployeeDataEntryByDateRangeDataSet = new FindEmployeeDataEntryByDateRangeDataSet();
        FindDuplicateProjectMatrixDataSet TheFindDuplicateProjectMatrixDataSet = new FindDuplicateProjectMatrixDataSet();
        FindProjectMatrixByCustomerAssignedIDForEmailDataSet TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet = new FindProjectMatrixByCustomerAssignedIDForEmailDataSet();
        FindOverdueOpenProductionProjectsDataSet TheFindOverDueProductionProjectsDataSet = new FindOverdueOpenProductionProjectsDataSet();
        FindEmployeeByEmployeeIDDataSet TheFindEmployeeByEmployeeIDDataSet = new FindEmployeeByEmployeeIDDataSet();
        FindGEOFenceTransactionDateRangeDataSet TheFindGeoFenceTransactionDateRangeDataSet = new FindGEOFenceTransactionDateRangeDataSet();

        FindDuplicateVehicleAssignmentsDataSet aFindDuplicateVehicleAsignmentsDataSet;
        FindDuplicateVehicleAssignmentsDataSet TheFindDuplicateVehicleAssignmentsDataSet = new FindDuplicateVehicleAssignmentsDataSet();
        FindDuplicateVehicleAssignmentsDataSetTableAdapters.FindDuplicateVehicleAssignmentsTableAdapter aFindDuplicateVehicleAssignmentsTableAdapter;

        int gintLogCounter;
        int gintLogUpperLimit;
        bool gblnItemRan = false;

        int gintNumberOfRecords;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            TheMessagesClass.CloseTheProgram();
        }
        private FindDuplicateVehicleAssignmentsDataSet FindDuplicateVehicleAssignments()
        {
            try
            {
                aFindDuplicateVehicleAsignmentsDataSet = new FindDuplicateVehicleAssignmentsDataSet();
                aFindDuplicateVehicleAssignmentsTableAdapter = new FindDuplicateVehicleAssignmentsDataSetTableAdapters.FindDuplicateVehicleAssignmentsTableAdapter();
                aFindDuplicateVehicleAssignmentsTableAdapter.Fill(aFindDuplicateVehicleAsignmentsDataSet.FindDuplicateVehicleAssignments);
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Main Window // Find Duplicate Vehicle Assignments " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }


            return aFindDuplicateVehicleAsignmentsDataSet;
        }
        private void UpdateGrid()
        {
            //setting variables
            int intNumberOfRecords;
            bool blnFatalError = false;
            DateTime datEndDate = DateTime.Now;
            DateTime datStartDate = DateTime.Now;

            try
            {
                TheFindDuplicateVehicleAssignmentsDataSet = FindDuplicateVehicleAssignments();

                intNumberOfRecords = TheFindDuplicateVehicleAssignmentsDataSet.FindDuplicateVehicleAssignments.Rows.Count;

                if(intNumberOfRecords > 0)
                {
                    blnFatalError = TheSendEmailClass.SendEmail("tholmes@bluejaycommunications.com", "New Event Log Entry", "We Have a Duplicate Vehicle Assignment Event");
                    if (blnFatalError == true)
                        throw new Exception();

                    TheMessagesClass.ErrorMessage("We Have a Duplicate Vehicle Assignment Event");
                }

                TheFindDuplicateProjectMatrixDataSet = TheProjectMatrixClass.FindDuplicateProjectMatrix();

                intNumberOfRecords = TheFindDuplicateProjectMatrixDataSet.FindDuplicateProjectMatrix.Rows.Count;

                if(intNumberOfRecords > 0)
                {
                    blnFatalError = CreateDuplicateProjectEmail();

                    if (blnFatalError == true)
                        throw new Exception();
                }

                datEndDate = TheDateSearchClass.RemoveTime(datEndDate);

                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 5);

                datEndDate = TheDateSearchClass.AddingDays(datEndDate, 1);

                TheFindEventLogByDateRangeDataSet = TheEventLogClass.FindEventLogByDateRange(datStartDate, datEndDate);

                intNumberOfRecords = TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange.Rows.Count;

                if (intNumberOfRecords != gintNumberOfRecords)
                {
                    gintNumberOfRecords = intNumberOfRecords;
                    intNumberOfRecords -= 1;

                    blnFatalError = TheSendEmailClass.SendEmail("tholmes@bluejaycommunications.com", "New Event Log Entry", TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange[0].LogEntry);

                    if (blnFatalError == true)
                        throw new Exception();

                    blnFatalError = TheSendEmailClass.SendEmail("mharmon@bluejaycommunications.com", "New Event Log Entry", TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange[0].LogEntry);

                    if (blnFatalError == true)
                        throw new Exception();

                    dgrResults.ItemsSource = TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange;

                }

                dgrResults.ItemsSource = TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Main Window // Update Grid " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }           
            
        }
        private bool CreateDuplicateProjectEmail()
        {
            int intNumberOfRecords;
            int intCounter;
            string strHeader = "DUPLICATE PROJECTS";
            string strMessage;
            bool blnFatalError = false;
            int intSecondCounter;
            int intSecondNumberOfRecords;
            string strCustomerID;
            string strEmailAddress = "";
            string strEmployeeName;

            try
            {
                //loading up the data
                intNumberOfRecords = TheFindDuplicateProjectMatrixDataSet.FindDuplicateProjectMatrix.Rows.Count - 1;

                strMessage = "<h1>" + strHeader + " - Do Not Reply</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td>TransactionID</td>";
                strMessage += "<td>Transaction Date</td>";
                strMessage += "<td>Our Project ID</td>";
                strMessage += "<td>Custoner Project ID</td>";
                strMessage += "<td>Project Name</td>";
                strMessage += "<td>Employee Name</td>";
                strMessage += "</tr>";
                strMessage += "<p>          </p>";

                if (intNumberOfRecords > -1)
                {
                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        strCustomerID = TheFindDuplicateProjectMatrixDataSet.FindDuplicateProjectMatrix[intCounter].CustomerAssignedID;

                        TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet = TheProjectMatrixClass.FindProjectMatrixByCustomerAssignedIDForEmail(strCustomerID);

                        intSecondNumberOfRecords = TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet.FindProjectMatrixByCustomerAssignedIDForEmail.Rows.Count;

                        if(intSecondNumberOfRecords > 0)
                        {
                            for(intSecondCounter = 0; intSecondCounter < intSecondNumberOfRecords; intSecondCounter++)
                            {
                                strEmployeeName = TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet.FindProjectMatrixByCustomerAssignedIDForEmail[intSecondCounter].FirstName + " ";
                                strEmployeeName += TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet.FindProjectMatrixByCustomerAssignedIDForEmail[intSecondCounter].LastName;

                                strMessage += "<tr>";
                                strMessage += "<td>" + Convert.ToString(TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet.FindProjectMatrixByCustomerAssignedIDForEmail[intSecondCounter].TransactionID) + "</td>";
                                strMessage += "<td>" + Convert.ToString(TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet.FindProjectMatrixByCustomerAssignedIDForEmail[intSecondCounter].TransactionDate) + "</td>";
                                strMessage += "<td>" + TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet.FindProjectMatrixByCustomerAssignedIDForEmail[intSecondCounter].AssignedProjectID + "</td>";
                                strMessage += "<td>" + TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet.FindProjectMatrixByCustomerAssignedIDForEmail[intSecondCounter].CustomerAssignedID + "</td>";
                                strMessage += "<td>" + TheFindProjectMatrixByCustomerAssignedIDForEmailDataSet.FindProjectMatrixByCustomerAssignedIDForEmail[intSecondCounter].ProjectName + "</td>";
                                strMessage += "<td>" + strEmployeeName + "</td>";
                                strMessage += "</tr>";
                            }
                        }                        
                    }
                }

                strMessage += "</table>";

                blnFatalError = TheSendEmailClass.SendEmail("tholmes@bluejaycommunications.com", strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = TheSendEmailClass.SendEmail("mharmon@bluejaycommunications.com", strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = TheSendEmailClass.SendEmail("csimmons@bluejaycommunications.com", strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();

                blnFatalError = TheSendEmailClass.SendEmail("mchapman@bluejaycommunications.com", strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Main Window // Data Entry Reports " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());                
            }

            return blnFatalError;
        }
        private void BeginTheProcess(object sender, EventArgs e)
        {
            UpdateGrid();
            UpdateVehicleStatus();
            SendVehicleReports();
            CheckEmployeePayRate();            
        }
        private void ChangeVehicleInYardToWarehouse()
        {
            DateTime datStartDate = DateTime.Now;
            DateTime datEndDate = DateTime.Now;
            int intCounter;
            int intNumberOfRecords;
            int intVehicleID;
            int intWarehouseID;
            int intTransactionID;
            string strWarehouse;
            bool blnFatalError = false;

            try
            {
                datStartDate = TheDateSearchClass.RemoveTime(datStartDate);
                datEndDate = TheDateSearchClass.AddingDays(datStartDate, 1);

                TheFindVehiclesInYardShowingVehicleIDDateRangeDataSet = TheVehicleInYardClass.FindVehiclesInYardShowingVehicleIDDateRange(datStartDate, datEndDate);

                intNumberOfRecords = TheFindVehiclesInYardShowingVehicleIDDateRangeDataSet.FindVehiclesInYardShowingVehicleDateRange.Rows.Count - 1;

                if(intNumberOfRecords > -1)
                {
                    for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        intVehicleID = TheFindVehiclesInYardShowingVehicleIDDateRangeDataSet.FindVehiclesInYardShowingVehicleDateRange[intCounter].VehicleID;
                        strWarehouse = TheFindVehiclesInYardShowingVehicleIDDateRangeDataSet.FindVehiclesInYardShowingVehicleDateRange[intCounter].AssignedOffice;

                        TheFindWarehouseByWarehouseNameDataSet = TheEmployeeClass.FindWarehouseByWarehouseName(strWarehouse);

                        intWarehouseID = TheFindWarehouseByWarehouseNameDataSet.FindWarehouseByWarehouseName[0].EmployeeID;

                        TheFindCurrentAssignedVehicleMainByVehicleIDDataSet = TheVehicleAssignmentClass.FindCurrentAssignedVehicleMainByVehicleID(intVehicleID);

                        intTransactionID = TheFindCurrentAssignedVehicleMainByVehicleIDDataSet.FindCurrentAssignedVehicleMainByVehicleID[0].TransactionID;

                        blnFatalError = TheVehicleAssignmentClass.UpdateCurrentVehicleAssignment(intTransactionID, false);

                        if (blnFatalError == true)
                            throw new Exception();

                        blnFatalError = TheVehicleAssignmentClass.InsertVehicleAssignment(intVehicleID, intWarehouseID);

                        if (blnFatalError == true)
                            throw new Exception();

                        blnFatalError = TheVehicleMainClass.UpdateVehicleMainEmployeeID(intVehicleID, intWarehouseID);

                        if (blnFatalError == true)
                            throw new Exception();

                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Main Window // Change Vehicle In Yard To Warehouse " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void CheckEmployeePayRate()
        {
            int intCounter;
            int intNumberOfRecords;
            int intEmployeeID;
            int intRecordsReturned;
            bool blnFatalError = false;

            try
            {
                TheFindActiveEmployeesDataSet = TheEmployeeClass.FindActiveEmployees();

                intNumberOfRecords = TheFindActiveEmployeesDataSet.FindActiveEmployees.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intEmployeeID = TheFindActiveEmployeesDataSet.FindActiveEmployees[intCounter].EmployeeID;

                    TheFindEmployeeLaborRateDataSet = TheEmployeeLaborRateClass.FindEmployeeLaborRate(intEmployeeID);

                    intRecordsReturned = TheFindEmployeeLaborRateDataSet.FindEmployeeLaborRate.Rows.Count;

                    if(intRecordsReturned == 0)
                    {
                        blnFatalError = TheEmployeeLaborRateClass.InsertEmployeeLaborRate(intEmployeeID, 1);

                        if (blnFatalError == true)
                            throw new Exception();
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Cheeck Employee Pay Rate " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

            
            
        }
        private void SendVehicleReports()
        {
            DateTime datTodaysDate = DateTime.Now;
            DateTime datTransactionDate;
            DateTime datStartDate;
            DateTime datEndDate;

            try
            {
                TheVehicleExceptionEmailDataSet = TheVehicleExceptionEmailClass.GetVehicleExceptionEmailInfo();
                TheWeeklyVehicleReportsDateDataSet = TheVehicleExceptionEmailClass.GetWeeklyVehicleReportsDateInfo();

                datTransactionDate = TheVehicleExceptionEmailDataSet.vehicleexceptionemail[0].TransactionDate;

                datTransactionDate = TheDateSearchClass.AddingDays(datTransactionDate, 1);
                datStartDate = TheDateSearchClass.SubtractingDays(datTodaysDate, 1);
                datStartDate = TheDateSearchClass.RemoveTime(datStartDate);

                if (datTodaysDate > datTransactionDate)
                {
                    if (datTransactionDate.DayOfWeek == DayOfWeek.Saturday)
                    {
                        TheVehicleExceptionEmailDataSet.vehicleexceptionemail[0].TransactionDate = datTransactionDate;
                        TheVehicleExceptionEmailClass.UpdateVehicleExceptionEmailDB(TheVehicleExceptionEmailDataSet);
                    }
                    else if (datTransactionDate.DayOfWeek == DayOfWeek.Sunday)
                    {
                        TheVehicleExceptionEmailDataSet.vehicleexceptionemail[0].TransactionDate = datTransactionDate;
                        TheVehicleExceptionEmailClass.UpdateVehicleExceptionEmailDB(TheVehicleExceptionEmailDataSet);
                    }
                    else
                    {
                        ChangeVehicleInYardToWarehouse();
                        datTodaysDate = TheDateSearchClass.RemoveTime(datTodaysDate);
                        datEndDate = TheDateSearchClass.AddingDays(datTodaysDate, 1);
                        TheVehicleExceptionEmailDataSet.vehicleexceptionemail[0].TransactionDate = datTransactionDate;
                        TheAutomatedVehicleReportsClass.RunAutomatedReports(datTodaysDate);
                        TheVehicleExceptionEmailClass.UpdateVehicleExceptionEmailDB(TheVehicleExceptionEmailDataSet);
                        TheUpdatingWorkTaskStatsClass.UpdateWorkTaskStatsTable();
                        DataEntryReports(datStartDate, datEndDate);
                        SendOverdueProjectReport();
                        //SendVehicleAfterHourActivity();
                    }
                }

                if(datTodaysDate.DayOfWeek == DayOfWeek.Monday)
                {
                    datTodaysDate = TheDateSearchClass.RemoveTime(datTodaysDate);

                     if(datTodaysDate > TheWeeklyVehicleReportsDateDataSet.weeklyvehiclereportsdate[0].LastWeeklyReport)
                     {
                        TheAutomatedVehicleReportsClass.RunWeeklyVehicleInspectionReport();

                        TheAutomatedVehicleReportsClass.RunWeeklyVehiclesInYardReport();

                        TheWeeklyVehicleReportsDateDataSet.weeklyvehiclereportsdate[0].LastWeeklyReport = datTodaysDate;

                        TheVehicleExceptionEmailClass.UpdateWeeklyVehicleReportsDB(TheWeeklyVehicleReportsDateDataSet);

                        TheAutomatedProductioinReportsClass.RunAutomatedProductionReports();
                     }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay ERP // Main Window // Begin The Process " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void SendVehicleAfterHourActivity()
        {
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress = "itadmin@bluejaycommunications.com";
            DateTime datStartDate = DateTime.Now;
            DateTime datEndDate = DateTime.Now;
            string strEventDate = "";
            string strVehicleNumber;
            string strEmployee;
            string strAssignedOffice;
            string strDriver;
            bool blnInSide;
            string strInOut = "";
            string strHeader = "BJC Vehicle After Hour Report";
            string strMessage;
            bool blnFatalError;

            try
            {
                datStartDate = TheDateSearchClass.RemoveTime(datStartDate);
                datEndDate = TheDateSearchClass.RemoveTime(datEndDate);
                datStartDate = TheDateSearchClass.SubtractingDays(datStartDate, 1);
                datStartDate = datStartDate.AddHours(18);
                datEndDate = datEndDate.AddHours(6);

                TheFindGeoFenceTransactionDateRangeDataSet = TheGeoFenceClass.FindGEOFenceTransactionDAteRanage(datStartDate, datEndDate);

                strMessage = "<h1>" + strHeader + " - Do Not Reply</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td>Event Date</td>";
                strMessage += "<td>Vehicle Number</td>";
                strMessage += "<td>Employee</td>";
                strMessage += "<td>Office</td>";
                strMessage += "<td>Driver</td>";
                strMessage += "<td>In Or Out</td>";
                strMessage += "</tr>";
                strMessage += "<p>          </p>";

                intNumberOfRecords = TheFindGeoFenceTransactionDateRangeDataSet.FindGEOFenceTransasctionDateRange.Rows.Count;

                if (intNumberOfRecords > 0)
                {
                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strEventDate = Convert.ToString(TheFindGeoFenceTransactionDateRangeDataSet.FindGEOFenceTransasctionDateRange[intCounter].EventDate);
                        strVehicleNumber = TheFindGeoFenceTransactionDateRangeDataSet.FindGEOFenceTransasctionDateRange[intCounter].VehicleNumber;
                        strEmployee = TheFindGeoFenceTransactionDateRangeDataSet.FindGEOFenceTransasctionDateRange[intCounter].FirstName + " ";
                        strEmployee += TheFindGeoFenceTransactionDateRangeDataSet.FindGEOFenceTransasctionDateRange[intCounter].LastName;
                        strAssignedOffice = TheFindGeoFenceTransactionDateRangeDataSet.FindGEOFenceTransasctionDateRange[intCounter].AssignedOffice;
                        strDriver = TheFindGeoFenceTransactionDateRangeDataSet.FindGEOFenceTransasctionDateRange[intCounter].Driver;
                        blnInSide = TheFindGeoFenceTransactionDateRangeDataSet.FindGEOFenceTransasctionDateRange[intCounter].InSide;

                        if (blnInSide == true)
                            strInOut = "IN";
                        else if (blnInSide == false)
                            strInOut = "OUT";

                        strMessage += "<tr>";
                        strMessage += "<td>" + strEventDate + "</td>";
                        strMessage += "<td>" + strVehicleNumber + "</td>";
                        strMessage += "<td>" + strEmployee + "</td>";
                        strMessage += "<td>" + strAssignedOffice + "</td>";
                        strMessage += "<td>" + strDriver+ "</td>";
                        strMessage += "<td>" + strInOut + "</td>";
                        strMessage += "</tr>";
                    }
                }

                strMessage += "</table>";

                blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Send Overdue Project Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void SendOverdueProjectReport()
        {
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress = "erpprojectreports@bluejaycommunications.com";
            DateTime datTransactionDate = DateTime.Now;
            string strCustomer;
            string strCustomerAssignedID;
            string strAssignedProjectID;
            string strProjectName;
            string strDateReceived;
            string strECDDate;
            string strAssignedOffice;
            string strStatus;
            int intWarehouseID;
            string strHeader = "BJC Overdue Projects Report";
            string strMessage;
            bool blnFatalError;

            try
            {
                datTransactionDate = TheDateSearchClass.AddingDays(datTransactionDate, 3);

                TheFindOverDueProductionProjectsDataSet = TheProductionProjectClass.FindOverdueProductionProjects(datTransactionDate);

                strMessage = "<h1>" + strHeader + " - Do Not Reply</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td>Customer</td>";
                strMessage += "<td>Custoner Project ID</td>";
                strMessage += "<td>BJC Project ID</td>";
                strMessage += "<td>Project Name</td>";
                strMessage += "<td>Date Received</td>";
                strMessage += "<td>ECD Date</td>";
                strMessage += "<td>Assigned Office</td>";
                strMessage += "<td>Status</td>";
                strMessage += "</tr>";
                strMessage += "<p>          </p>";

                intNumberOfRecords = TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects.Rows.Count;

                if(intNumberOfRecords > 0)
                {
                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strCustomer = TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects[intCounter].Customer;
                        strCustomerAssignedID = TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects[intCounter].CustomerAssignedID;
                        strAssignedProjectID = TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects[intCounter].AssignedProjectID;
                        strProjectName = TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects[intCounter].ProjectName;
                        strDateReceived = Convert.ToString(TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects[intCounter].DateReceived);
                        strECDDate = Convert.ToString(TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects[intCounter].ECDDate);
                        strStatus = TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects[intCounter].WorkOrderStatus;

                        intWarehouseID = TheFindOverDueProductionProjectsDataSet.FindOverdueOpenProductionProjects[intCounter].AssignedOfficeID;

                        TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intWarehouseID);

                        strAssignedOffice = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName;

                        strMessage += "<tr>";
                        strMessage += "<td>" + strCustomer + "</td>";
                        strMessage += "<td>" + strCustomerAssignedID + "</td>";
                        strMessage += "<td>" + strAssignedProjectID + "</td>";
                        strMessage += "<td>" + strProjectName + "</td>";
                        strMessage += "<td>" + strDateReceived + "</td>";
                        strMessage += "<td>" + strECDDate + "</td>";
                        strMessage += "<td>" + strAssignedOffice + "</td>";
                        strMessage += "<td>" + strStatus + "</td>";
                        strMessage += "</tr>";
                    }
                }

                strMessage += "</table>";

                blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Send Overdue Project Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

        }
        private void DataEntryReports(DateTime datStartDate, DateTime datEndDate)
        {
            //setting up the variables
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress = "itadmin@bluejaycommunications.com";
            string strHeader = "Employee ERP Logins";
            string strMessage = "";
            bool blnFatalError = false;

            try
            {
                //loading up the data
                TheFindEmployeeDataEntryByDateRangeDataSet = TheEmployeeDataEntryClass.FindEmployeeDataEntryByDateRange(datStartDate, datEndDate);

                intNumberOfRecords = TheFindEmployeeDataEntryByDateRangeDataSet.FindEmployeeDateEntryByDateRange.Rows.Count - 1;

                strMessage = "<h1>" + strHeader + " - Do Not Reply</h1>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td>Transaction Date</td>";
                strMessage += "<td>First Name</td>";
                strMessage += "<td>Last Name</td>";
                strMessage += "<td>Home Office</td>";
                strMessage += "<td>Window Entered</td>";
                strMessage += "</tr>";
                strMessage += "<p>          </p>";

                if(intNumberOfRecords > -1)
                {
                    for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        strMessage += "<tr>";
                        strMessage += "<td>" + Convert.ToString(TheFindEmployeeDataEntryByDateRangeDataSet.FindEmployeeDateEntryByDateRange[intCounter].TransactionDate) + "</td>";
                        strMessage += "<td>" + TheFindEmployeeDataEntryByDateRangeDataSet.FindEmployeeDateEntryByDateRange[intCounter].FirstName + "</td>";
                        strMessage += "<td>" + TheFindEmployeeDataEntryByDateRangeDataSet.FindEmployeeDateEntryByDateRange[intCounter].LastName + "</td>";
                        strMessage += "<td>" + TheFindEmployeeDataEntryByDateRangeDataSet.FindEmployeeDateEntryByDateRange[intCounter].HomeOffice + "</td>";
                        strMessage += "<td>" + TheFindEmployeeDataEntryByDateRangeDataSet.FindEmployeeDateEntryByDateRange[intCounter].WindowEntered + "</td>";
                        strMessage += "</tr>";
                    }
                }

                strMessage += "</table>";

                blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Main Window // Data Entry Reports " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void UpdateVehicleStatus()
        {
            int intCounter;
            int intNumberOfRecords;
            int intVehicleID;
            int intRecordsReturned;
            bool blnFatalError = false;
            string strStatus;

            try
            {
                //getting the vehicles
                TheFindActiveVehiclesDataSet = TheVehicleClass.FindActiveVehicles();

                intNumberOfRecords = TheFindActiveVehiclesDataSet.FindActiveVehicles.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intVehicleID = TheFindActiveVehiclesDataSet.FindActiveVehicles[intCounter].VehicleID;

                    TheFindOpenVehicleProblemsByVehicleIDDataSet = TheVehicleProblemClass.FindOpenVehicleProblemsbyVehicleID(intVehicleID);
                    
                    intRecordsReturned = TheFindOpenVehicleProblemsByVehicleIDDataSet.FindOpenVehicleProblemsByVehicleID.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        //checking to see what the status is
                        TheFindVehicleStatusByVehicleIDDataSet = TheVehicleStatusClass.FindVehicleStatusByVehicleID(intVehicleID);

                        intRecordsReturned = TheFindVehicleStatusByVehicleIDDataSet.FindVehicleStatusByVehicleID.Rows.Count;
                        
                        if(intRecordsReturned == 1)
                        {
                            strStatus = TheFindVehicleStatusByVehicleIDDataSet.FindVehicleStatusByVehicleID[0].VehicleStatus;

                            if((strStatus == "AVAILABLE") || (strStatus == "SIGNED OUT"))
                            {
                                blnFatalError = TheVehicleStatusClass.UpdateVehicleStatus(intVehicleID, "NEEDS WORK", DateTime.Now);

                                if (blnFatalError == true)
                                    throw new Exception();
                            }
                            if(strStatus == "NO PROBLEM")
                            {
                                blnFatalError = TheVehicleStatusClass.UpdateVehicleStatus(intVehicleID, "NEEDS WORK", DateTime.Now);

                                if (blnFatalError == true)
                                    throw new Exception();
                            }
                        }
                    }
                    else if (intRecordsReturned == 0)
                    {
                        TheFindVehicleStatusByVehicleIDDataSet = TheVehicleStatusClass.FindVehicleStatusByVehicleID(intVehicleID);

                        intRecordsReturned = TheFindVehicleStatusByVehicleIDDataSet.FindVehicleStatusByVehicleID.Rows.Count;

                        if(intRecordsReturned == 1)
                        {
                            if(TheFindVehicleStatusByVehicleIDDataSet.FindVehicleStatusByVehicleID[0].VehicleStatus != "NO PROBLEM")
                            {
                                blnFatalError = TheVehicleStatusClass.UpdateVehicleStatus(intVehicleID, "NO PROBLEM", DateTime.Now);

                                if (blnFatalError == true)
                                    throw new Exception();
                            }
                        }
                        else if (intRecordsReturned == 0)
                        {
                            blnFatalError = TheVehicleStatusClass.InsertVehicleStatus(intVehicleID, "NO PROBLEM", DateTime.Now);

                            if (blnFatalError == true)
                                throw new Exception();
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Update Vehicle Status " + Ex.Message);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            gintNumberOfRecords = 0;
            DateTime datEndDate = DateTime.Now;
            DateTime datStartDate;
            
            datEndDate = TheDateSearchClass.RemoveTime(datEndDate);
            datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 5);
            datEndDate = TheDateSearchClass.AddingDays(datEndDate, 1);

            TheFindEventLogByDateRangeDataSet = TheEventLogClass.FindEventLogByDateRange(datStartDate, datEndDate);

            gintNumberOfRecords = TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange.Rows.Count;
            UpdateGrid();
                    
            MyTimer.Tick += new EventHandler(BeginTheProcess);
            MyTimer.Interval = new TimeSpan(0, 0, 30);
            MyTimer.Start();
            
        }
       
    }
}
