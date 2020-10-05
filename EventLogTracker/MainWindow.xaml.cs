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
using System.Data;

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
                        //ChangeVehicleInYardToWarehouse();
                        datTodaysDate = TheDateSearchClass.RemoveTime(datTodaysDate);
                        datEndDate = TheDateSearchClass.AddingDays(datTodaysDate, 1);
                        TheVehicleExceptionEmailDataSet.vehicleexceptionemail[0].TransactionDate = datTransactionDate;
                        TheAutomatedVehicleReportsClass.RunAutomatedReports(datTodaysDate);
                        TheVehicleExceptionEmailClass.UpdateVehicleExceptionEmailDB(TheVehicleExceptionEmailDataSet);
                        //TheUpdatingWorkTaskStatsClass.UpdateWorkTaskStatsTable();
                        DataEntryReports(datStartDate, datEndDate);
                    }
                }

                if(datTodaysDate.DayOfWeek == DayOfWeek.Monday)
                {
                    datTodaysDate = TheDateSearchClass.RemoveTime(datTodaysDate);

                     if(datTodaysDate > TheWeeklyVehicleReportsDateDataSet.weeklyvehiclereportsdate[0].LastWeeklyReport)
                    {
                        TheAutomatedVehicleReportsClass.RunWeeklyVehicleInspectionReport();

                        //TheAutomatedVehicleReportsClass.RunWeeklyVehiclesInYardReport();

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
