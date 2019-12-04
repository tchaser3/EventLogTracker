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
        private void UpdateGrid()
        {
            //setting variables
            int intNumberOfRecords;
            bool blnEmailSent;
            DateTime datEndDate = DateTime.Now;
            DateTime datStartDate = DateTime.Now;

            try
            {
                datEndDate = TheDateSearchClass.RemoveTime(datEndDate);

                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 5);

                datEndDate = TheDateSearchClass.AddingDays(datEndDate, 1);

                TheFindEventLogByDateRangeDataSet = TheEventLogClass.FindEventLogByDateRange(datStartDate, datEndDate);

                intNumberOfRecords = TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange.Rows.Count;

                if (intNumberOfRecords != gintNumberOfRecords)
                {
                    gintNumberOfRecords = intNumberOfRecords;
                    intNumberOfRecords -= 1;

                    blnEmailSent = TheSendEmailClass.SendEmail("tholmes@bluejaycommunications.com", "New Event Log Entry", TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange[0].LogEntry);

                    blnEmailSent = TheSendEmailClass.SendEmail("mharmon@bluejaycommunications.com", "New Event Log Entry", TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange[0].LogEntry);


                    dgrResults.ItemsSource = TheFindEventLogByDateRangeDataSet.FindEventLogEntriesByDateRange;

                    if (blnEmailSent == false)
                    {
                        TheMessagesClass.ErrorMessage("Email Was Not Sent");
                    }
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
            ChangeVehicleInYardToWarehouse();
            TheUpdatingWorkTaskStatsClass.UpdateWorkTaskStatsTable();
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

            try
            {
                TheVehicleExceptionEmailDataSet = TheVehicleExceptionEmailClass.GetVehicleExceptionEmailInfo();
                TheWeeklyVehicleReportsDateDataSet = TheVehicleExceptionEmailClass.GetWeeklyVehicleReportsDateInfo();

                datTransactionDate = TheVehicleExceptionEmailDataSet.vehicleexceptionemail[0].TransactionDate;

                datTransactionDate = TheDateSearchClass.AddingDays(datTransactionDate, 1);

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
                        datTodaysDate = TheDateSearchClass.RemoveTime(datTodaysDate);
                        TheVehicleExceptionEmailDataSet.vehicleexceptionemail[0].TransactionDate = datTransactionDate;
                        TheAutomatedVehicleReportsClass.RunAutomatedReports(datTodaysDate);
                        TheVehicleExceptionEmailClass.UpdateVehicleExceptionEmailDB(TheVehicleExceptionEmailDataSet);
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
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Blue Jay ERP // Main Window // Begin The Process " + Ex.Message);

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
            //LoadEventLogTable();
                    
            MyTimer.Tick += new EventHandler(BeginTheProcess);
            MyTimer.Interval = new TimeSpan(0, 0, 30);
            MyTimer.Start();
            
        }
       
    }
}
