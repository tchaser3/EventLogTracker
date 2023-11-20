using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NewEventLogDLL;
using EmployeeProjectAssignmentDLL;
using EmployeePunchedHoursDLL;
using NewEmployeeDLL;
using DateSearchDLL;
using DataValidationDLL;
using Microsoft.Win32;
using DesignProductivityDLL;
using AfterHoursWorkDLL;

namespace EventLogTracker
{
    class RunPunchedVsProduction
    {
        //Setting up classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeProjectAssignmentClass TheEmployeeProjectAssignmentClass = new EmployeeProjectAssignmentClass();
        EmployeePunchedHoursClass TheEmployeePunchedHoursClass = new EmployeePunchedHoursClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        DesignProductivityClass TheDesignProductivityClass = new DesignProductivityClass();
        SendEmailClass TheSendEmailClass = new SendEmailClass();
        AfterHoursWorkClass TheAfterHoursWorkClass = new AfterHoursWorkClass();

        EmployeeProductionPunchDataSet TheEmployeeProductionPunchDataSet = new EmployeeProductionPunchDataSet();
        FindActiveNonExemptEmployeesByPayDateDataSet TheFindActiveNoExemptEmployeesDataSet = new FindActiveNonExemptEmployeesByPayDateDataSet();
        FindEmployeeByEmployeeIDDataSet TheFindEmployeeByEmployeeIDDataSet = new FindEmployeeByEmployeeIDDataSet();
        FindEmployeeProductionHoursOverPayPeriodDataSet TheFindEmployeeProductionHoursOverPayPeriodDataSet = new FindEmployeeProductionHoursOverPayPeriodDataSet();
        FindEmployeePunchedHoursDataSet TheFindEmployeePunchedHoursDataSet = new FindEmployeePunchedHoursDataSet();
        ProductionManagerStatsDataSet TheProductionManagerStatsDataSet = new ProductionManagerStatsDataSet();
        ProductionManagerStatsDataSet TheFinalProductionManagerStatsDataSet = new ProductionManagerStatsDataSet();
        FindDesignTotalEmployeeProductivityHoursDataSet TheFindDesignTotalEmployeeProductivityHoursDataSet = new FindDesignTotalEmployeeProductivityHoursDataSet();
        EmployeeOvertimeDataSet TheEmployeeOvertimeDataSet = new EmployeeOvertimeDataSet();
        FindEmployeesOverFortyHoursDataSet TheFindEmployeesOverFortyHoursDataSet = new FindEmployeesOverFortyHoursDataSet();
        FindEmployeeProductionForMiscCodeDataSet TheFindEmployeeProductionForMiscCodeDataSet = new FindEmployeeProductionForMiscCodeDataSet();
        FindEmployeeProductivityMiscTotalHoursDataSet TheFindEmployeeProductivityMiscTotalHoursDataSet = new FindEmployeeProductivityMiscTotalHoursDataSet();
        FindEmployeeAfterHoursWorkThiryDayReportDataSet TheFindEmployeeAfterHoursWorkThiryDayReportDataSet = new FindEmployeeAfterHoursWorkThiryDayReportDataSet();
        EmployeeAfterHourWorkReportDataSet TheEmployeeAfterHourWorkReportDataSet = new EmployeeAfterHourWorkReportDataSet();

        string[] gstrLastName = new string[20];

        int gintManagerCounter;
        int gintManagerUpperLimit;
        decimal gdecTotalEntries;
        decimal gdecTotalAcceptable;
        decimal gdecStandardDeviation;
        decimal gdecMean;

        public void RunPunchedVSProductionReport()
        {
            //setting local variables
            int intCounter;
            int intNumberOfRecords;
            int intEmployeeID;
            int intManagerID;
            string strFirstName;
            string strLastName;
            string strManagerFirstName;
            string strManagerLastName;
            int intRecordReturned;
            bool blnFatalError = false;
            decimal decHoursPunched;
            decimal decProductiveHours;
            int intManagerCounter;
            bool blnItemFound;
            decimal decVariance;
            int intTotalNumber;
            int intSelectedIndex;
            int intTotalNumberNeeded;
            decimal decTopPercentage;
            decimal decWorkingPercentage;
            int intArrayCounter;
            int intArrayNumberOfRecords = -1;
            int intEmployeeCounter;
            int intNumberOfEmployees;
            decimal decCalVariance;
            decimal decVarianceMean;
            decimal decSumOfDifferences;
            decimal decAverageOfDifferences;
            DateTime datStartDate;
            DateTime datEndDate;
            string strEmailAddress = "ERPProjectReports@bluejaycommunications.com";
            string strHeader;
            string strMessage;
            string strManager;
            decimal decTotalPercentage;

            try
            {
                TheEmployeeProductionPunchDataSet.employees.Rows.Clear();
                TheProductionManagerStatsDataSet.productionmanager.Rows.Clear();
                TheFinalProductionManagerStatsDataSet.productionmanager.Rows.Clear();

                datEndDate = DateTime.Now;
                datEndDate = TheDateSearchClass.RemoveTime(datEndDate);
                datEndDate = TheDateSearchClass.SubtractingDays(datEndDate, 9);
                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 6);

                TheFindActiveNoExemptEmployeesDataSet = TheEmployeeClass.FindActiveNonExemptEmployeesByPayDate(datStartDate);

                intNumberOfRecords = TheFindActiveNoExemptEmployeesDataSet.FindActiveNonExemptEmployeesByPayDate.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    //getting employee info
                    intEmployeeID = TheFindActiveNoExemptEmployeesDataSet.FindActiveNonExemptEmployeesByPayDate[intCounter].EmployeeID;
                    strFirstName = TheFindActiveNoExemptEmployeesDataSet.FindActiveNonExemptEmployeesByPayDate[intCounter].FirstName;
                    strLastName = TheFindActiveNoExemptEmployeesDataSet.FindActiveNonExemptEmployeesByPayDate[intCounter].LastName;
                    intManagerID = TheFindActiveNoExemptEmployeesDataSet.FindActiveNonExemptEmployeesByPayDate[intCounter].ManagerID;

                    TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);

                    strManagerFirstName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName;
                    strManagerLastName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;

                    //getting employee punched hours
                    TheFindEmployeePunchedHoursDataSet = TheEmployeePunchedHoursClass.FindEmployeePunchedHours(intEmployeeID, datStartDate, datEndDate);

                    intRecordReturned = TheFindEmployeePunchedHoursDataSet.FindEmployeePunchedHours.Rows.Count;

                    if (intRecordReturned == 0)
                    {
                        decHoursPunched = 0;
                    }
                    else
                    {
                        decHoursPunched = TheFindEmployeePunchedHoursDataSet.FindEmployeePunchedHours[0].PunchedHours;
                    }

                    //getting production hours
                    TheFindEmployeeProductionHoursOverPayPeriodDataSet = TheEmployeeProjectAssignmentClass.FindEmployeeProductionHoursOverPayPeriodDataSet(intEmployeeID, datStartDate, datEndDate);

                    intRecordReturned = TheFindEmployeeProductionHoursOverPayPeriodDataSet.FindEmployeeProductionHoursOverPayPeriod.Rows.Count;

                    if (intRecordReturned == 0)
                    {
                        TheFindDesignTotalEmployeeProductivityHoursDataSet = TheDesignProductivityClass.FindDesignTotalEmployeeProductivityHours(intEmployeeID, datStartDate, datEndDate);

                        intRecordReturned = TheFindDesignTotalEmployeeProductivityHoursDataSet.FindDesignTotalEmployeeProductivityHours.Rows.Count;

                        if (intRecordReturned == 0)
                        {
                            decProductiveHours = 0;
                        }
                        else
                        {
                            decProductiveHours = TheFindDesignTotalEmployeeProductivityHoursDataSet.FindDesignTotalEmployeeProductivityHours[0].TotalHours;
                        }
                    }
                    else
                    {
                        decProductiveHours = TheFindEmployeeProductionHoursOverPayPeriodDataSet.FindEmployeeProductionHoursOverPayPeriod[0].ProductionHours;

                        TheFindDesignTotalEmployeeProductivityHoursDataSet = TheDesignProductivityClass.FindDesignTotalEmployeeProductivityHours(intEmployeeID, datStartDate, datEndDate);

                        intRecordReturned = TheFindDesignTotalEmployeeProductivityHoursDataSet.FindDesignTotalEmployeeProductivityHours.Rows.Count;

                        if (intRecordReturned > 0)
                        {
                            decProductiveHours += TheFindDesignTotalEmployeeProductivityHoursDataSet.FindDesignTotalEmployeeProductivityHours[0].TotalHours;
                        }
                    }

                    //loading the dataset
                    EmployeeProductionPunchDataSet.employeesRow NewEmployeeRow = TheEmployeeProductionPunchDataSet.employees.NewemployeesRow();

                    NewEmployeeRow.HomeOffice = TheFindActiveNoExemptEmployeesDataSet.FindActiveNonExemptEmployeesByPayDate[intCounter].HomeOffice;
                    NewEmployeeRow.FirstName = strFirstName;
                    NewEmployeeRow.LastName = strLastName;
                    NewEmployeeRow.ManagerFirstName = strManagerFirstName;
                    NewEmployeeRow.ManagerLastName = strManagerLastName;
                    NewEmployeeRow.ProductionHours = decProductiveHours;
                    NewEmployeeRow.PunchedHours = decHoursPunched;
                    NewEmployeeRow.HourVariance = decProductiveHours - decHoursPunched;

                    TheEmployeeProductionPunchDataSet.employees.Rows.Add(NewEmployeeRow);
                }

                intNumberOfEmployees = TheEmployeeProductionPunchDataSet.employees.Rows.Count;
                decCalVariance = 0;

                for (intEmployeeCounter = 0; intEmployeeCounter < intNumberOfEmployees; intEmployeeCounter++)
                {
                    decCalVariance += TheEmployeeProductionPunchDataSet.employees[intEmployeeCounter].HourVariance;
                }

                decVarianceMean = decCalVariance / Convert.ToDecimal(intNumberOfEmployees);

                decVarianceMean = Math.Round(decVarianceMean, 4);

                gdecMean = decVarianceMean;
                decSumOfDifferences = 0;

                for (intEmployeeCounter = 0; intEmployeeCounter < intNumberOfEmployees; intEmployeeCounter++)
                {
                    decSumOfDifferences += Convert.ToDecimal(Math.Pow(Convert.ToDouble(TheEmployeeProductionPunchDataSet.employees[intEmployeeCounter].HourVariance - gdecMean), 2));
                }

                decAverageOfDifferences = decSumOfDifferences / Convert.ToDecimal(intNumberOfEmployees);

                gdecStandardDeviation = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(decAverageOfDifferences)));

                gdecStandardDeviation = Math.Round(gdecStandardDeviation, 4);

                intNumberOfRecords = TheEmployeeProductionPunchDataSet.employees.Rows.Count - 1;
                gintManagerCounter = 0;
                gdecTotalAcceptable = 0;
                gdecTotalEntries = 0;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strManagerFirstName = TheEmployeeProductionPunchDataSet.employees[intCounter].ManagerFirstName;
                    strManagerLastName = TheEmployeeProductionPunchDataSet.employees[intCounter].ManagerLastName;
                    decVariance = TheEmployeeProductionPunchDataSet.employees[intCounter].HourVariance;
                    blnItemFound = false;
                    intArrayNumberOfRecords = -1;

                    if (gintManagerCounter > 0)
                    {
                        for (intManagerCounter = 0; intManagerCounter <= gintManagerUpperLimit; intManagerCounter++)
                        {
                            if (strManagerFirstName == TheProductionManagerStatsDataSet.productionmanager[intManagerCounter].FirstName)
                            {
                                if (strManagerLastName == TheProductionManagerStatsDataSet.productionmanager[intManagerCounter].LastName)
                                {
                                    blnItemFound = true;

                                    if ((decVariance >= -5 && decVariance <= 5))
                                    {
                                        TheProductionManagerStatsDataSet.productionmanager[intManagerCounter].ZeroToFive++;
                                        gdecTotalEntries++;
                                        gdecTotalAcceptable++;
                                    }
                                    else if (((decVariance < -5) && (decVariance >= -10)) || ((decVariance > 5) && (decVariance <= 10)))
                                    {
                                        TheProductionManagerStatsDataSet.productionmanager[intManagerCounter].FiveToTen++;
                                        gdecTotalEntries++;
                                    }
                                    else if ((decVariance < -10) || (decVariance > 10))
                                    {
                                        TheProductionManagerStatsDataSet.productionmanager[intManagerCounter].Above10++;
                                        gdecTotalEntries++;
                                    }
                                }
                            }
                        }
                    }

                    if (blnItemFound == false)
                    {
                        ProductionManagerStatsDataSet.productionmanagerRow NewManagerRow = TheProductionManagerStatsDataSet.productionmanager.NewproductionmanagerRow();

                        NewManagerRow.FirstName = strManagerFirstName;
                        NewManagerRow.LastName = strManagerLastName;
                        NewManagerRow.PercentInTolerance = 0;

                        if (Convert.ToDouble(decVariance) > -5.01)
                        {
                            NewManagerRow.ZeroToFive = 1;
                            NewManagerRow.FiveToTen = 0;
                            NewManagerRow.Above10 = 0;
                            gdecTotalEntries++;
                            gdecTotalAcceptable++;
                        }
                        else if ((Convert.ToDouble(decVariance) < -5) && (Convert.ToDouble(decVariance) > -10.01))
                        {
                            NewManagerRow.ZeroToFive = 0;
                            NewManagerRow.FiveToTen = 1;
                            NewManagerRow.Above10 = 0;
                            gdecTotalEntries++;
                        }
                        else if (Convert.ToDouble(decVariance) < -10)
                        {
                            NewManagerRow.ZeroToFive = 0;
                            NewManagerRow.FiveToTen = 0;
                            NewManagerRow.Above10 = 1;
                            gdecTotalEntries++;
                        }

                        TheProductionManagerStatsDataSet.productionmanager.Rows.Add(NewManagerRow);
                        gintManagerUpperLimit = gintManagerCounter;
                        gintManagerCounter++;
                    }
                }

                intNumberOfRecords = TheProductionManagerStatsDataSet.productionmanager.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intTotalNumber = TheProductionManagerStatsDataSet.productionmanager[intCounter].ZeroToFive;
                    intTotalNumber += TheProductionManagerStatsDataSet.productionmanager[intCounter].FiveToTen;
                    intTotalNumber += TheProductionManagerStatsDataSet.productionmanager[intCounter].Above10;

                    TheProductionManagerStatsDataSet.productionmanager[intCounter].PercentInTolerance = (Convert.ToDecimal(TheProductionManagerStatsDataSet.productionmanager[intCounter].ZeroToFive) / intTotalNumber) * 100;

                    TheProductionManagerStatsDataSet.productionmanager[intCounter].PercentInTolerance = Math.Round(TheProductionManagerStatsDataSet.productionmanager[intCounter].PercentInTolerance, 2);
                }

                decTotalPercentage = ((Math.Round(((gdecTotalAcceptable / gdecTotalEntries) * 100), 4)));

                intTotalNumberNeeded = 0;
                decTopPercentage = 1000;
                decWorkingPercentage = 0;
                strLastName = "";
                intSelectedIndex = -1;

                while (intTotalNumberNeeded <= intNumberOfRecords)
                {
                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        if (TheProductionManagerStatsDataSet.productionmanager[intCounter].PercentInTolerance <= decTopPercentage)
                        {
                            if (TheProductionManagerStatsDataSet.productionmanager[intCounter].PercentInTolerance >= decWorkingPercentage)
                            {
                                blnItemFound = false;

                                if (intArrayNumberOfRecords > -1)
                                {
                                    for (intArrayCounter = 0; intArrayCounter <= intArrayNumberOfRecords; intArrayCounter++)
                                    {
                                        if (TheProductionManagerStatsDataSet.productionmanager[intCounter].LastName == gstrLastName[intArrayCounter])
                                        {
                                            blnItemFound = true;
                                        }

                                    }
                                }


                                if (blnItemFound == false)
                                {
                                    decWorkingPercentage = TheProductionManagerStatsDataSet.productionmanager[intCounter].PercentInTolerance;
                                    intSelectedIndex = intCounter;
                                }
                            }
                        }
                    }

                    ProductionManagerStatsDataSet.productionmanagerRow NewManagerRow = TheFinalProductionManagerStatsDataSet.productionmanager.NewproductionmanagerRow();

                    NewManagerRow.Above10 = TheProductionManagerStatsDataSet.productionmanager[intSelectedIndex].Above10;
                    NewManagerRow.FirstName = TheProductionManagerStatsDataSet.productionmanager[intSelectedIndex].FirstName;
                    NewManagerRow.LastName = TheProductionManagerStatsDataSet.productionmanager[intSelectedIndex].LastName;
                    NewManagerRow.PercentInTolerance = TheProductionManagerStatsDataSet.productionmanager[intSelectedIndex].PercentInTolerance;
                    NewManagerRow.FiveToTen = TheProductionManagerStatsDataSet.productionmanager[intSelectedIndex].FiveToTen;
                    NewManagerRow.ZeroToFive = TheProductionManagerStatsDataSet.productionmanager[intSelectedIndex].ZeroToFive;

                    TheFinalProductionManagerStatsDataSet.productionmanager.Rows.Add(NewManagerRow);
                    intTotalNumberNeeded++;
                    decTopPercentage = decWorkingPercentage;
                    decWorkingPercentage = 0;
                    intArrayNumberOfRecords++;
                    gstrLastName[intArrayNumberOfRecords] = TheProductionManagerStatsDataSet.productionmanager[intSelectedIndex].LastName;
                }

                strHeader = "Employee Punched Vs Production Report For " + Convert.ToString(datEndDate);

                intNumberOfRecords = TheEmployeeProductionPunchDataSet.employees.Rows.Count;

                strMessage = "<h1> Employee Punched Vs Production Report For " + Convert.ToString(datEndDate) + "</h1>";
                strMessage += "<p>               </p>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Home Office</b></td>";
                strMessage += "<td><b>Manager</b></td>";
                strMessage += "<td><b>Punched Hours</b></td>";
                strMessage += "<td><b>Production Hours</b></td>";
                strMessage += "<td><b>Hour Variance</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                if(intNumberOfRecords > 0)
                {
                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strManager = TheEmployeeProductionPunchDataSet.employees[intCounter].ManagerFirstName + " ";
                        strManager += TheEmployeeProductionPunchDataSet.employees[intCounter].ManagerLastName;

                        strMessage += "<tr>";
                        strMessage += "<td>" + TheEmployeeProductionPunchDataSet.employees[intCounter].FirstName + "</td>";
                        strMessage += "<td>" + TheEmployeeProductionPunchDataSet.employees[intCounter].LastName + "</td>";
                        strMessage += "<td>" + TheEmployeeProductionPunchDataSet.employees[intCounter].HomeOffice + "</td>";
                        strMessage += "<td>" + strManager + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheEmployeeProductionPunchDataSet.employees[intCounter].PunchedHours) +"</td>";
                        strMessage += "<td>" + Convert.ToString(TheEmployeeProductionPunchDataSet.employees[intCounter].ProductionHours) + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheEmployeeProductionPunchDataSet.employees[intCounter].HourVariance) + "</td>";
                        strMessage += "</tr>";
                    }
                }

                strMessage += "</table>";
                strMessage += "<p>               </p>";

                blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();

                strHeader = "Manager Punched Vs Productivity Report For " + Convert.ToString(datEndDate);

                strMessage = "<h1>Employee Punched Vs Production Report For " + Convert.ToString(datEndDate) + "</h1>";
                strMessage += "<p>               </p>";
                strMessage += "<h3>Company Average &emsp;       " + Convert.ToString(gdecMean) + "</h3>";
                strMessage += "<h3>Company Standard Deviation &emsp;         " + Convert.ToString(gdecStandardDeviation) + "</h3>";
                strMessage += "<h3>Company Percentage &emsp;     " + Convert.ToString(decTotalPercentage);
                strMessage += "<p>               </p>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Zero to Five</b></td>";
                strMessage += "<td><b>Five To Ten</b></td>";
                strMessage += "<td><b>Above Ten</b></td>";
                strMessage += "<td><b>Tolerance Percent</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                intNumberOfEmployees = TheFinalProductionManagerStatsDataSet.productionmanager.Rows.Count;

                if(intNumberOfEmployees > 0)
                {
                    for(intCounter = 0; intCounter < intNumberOfEmployees; intCounter++)
                    {
                        strMessage += "<tr>";
                        strMessage += "<td>" + TheFinalProductionManagerStatsDataSet.productionmanager[intCounter].FirstName + "</td>";
                        strMessage += "<td>" + TheFinalProductionManagerStatsDataSet.productionmanager[intCounter].LastName + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheFinalProductionManagerStatsDataSet.productionmanager[intCounter].ZeroToFive) + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheFinalProductionManagerStatsDataSet.productionmanager[intCounter].FiveToTen) + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheFinalProductionManagerStatsDataSet.productionmanager[intCounter].Above10) + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheFinalProductionManagerStatsDataSet.productionmanager[intCounter].PercentInTolerance) + "</td>";
                        strMessage += "</tr>";
                        strMessage += "<p>               </p>";
                    }
                }

                strMessage += "</table>";

                blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Run Punched Vs Production Class // Run Punched Vs Production Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void LoadAfterHoursDataSet()
        {
            int intCounter;
            int intNumberOfRecords;
            int intEmployeeID;
            string strManagerName = "";
            string strOfficeName = "";

            try
            {
                TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork.Rows.Clear();
                TheFindEmployeeAfterHoursWorkThiryDayReportDataSet = TheAfterHoursWorkClass.FindEmployeeOverNightWorkThirtyDayReport();

                intNumberOfRecords = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport.Rows.Count;

                if (intNumberOfRecords > 0)
                {
                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        intEmployeeID = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].ManagerID;

                        TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intEmployeeID);

                        if (TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID.Rows.Count > 0)
                        {
                            strManagerName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName + " ";
                            strManagerName += TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;
                        }

                        intEmployeeID = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].OfficeID;

                        TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intEmployeeID);

                        strOfficeName = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName;

                        EmployeeAfterHourWorkReportDataSet.employeeafterhourworkRow NewAfterWork = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork.NewemployeeafterhourworkRow();

                        NewAfterWork.AssignedOffice = strOfficeName;
                        NewAfterWork.BuildingAccess = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].BuildingAccess;
                        NewAfterWork.CustomerProjectID = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].CustomerAssignedID;
                        NewAfterWork.AssignedProjectID = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].AssignedProjectID;
                        NewAfterWork.Employee = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].Manager;
                        NewAfterWork.Manager = strManagerName;
                        NewAfterWork.OutTime = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].OutTime;
                        NewAfterWork.InETA = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].InETA;
                        NewAfterWork.ProjectName = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].ProjectName;
                        NewAfterWork.VehicleNumber = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].VehicleNumber;
                        NewAfterWork.WorkDate = TheFindEmployeeAfterHoursWorkThiryDayReportDataSet.FindEmployeeAfterHoursWorkThirtyDayReport[intCounter].WorkDate;

                        TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork.Rows.Add(NewAfterWork);
                    }
                }


            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "New Blue Jay ERP Browser //  After Hours Summary Report // Load Afterhours Dataset " + Ex.ToString());

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        public void EmailOverNightReport()
        {
            int intCounter;
            int intNumberOfRecords;
            string strWorkDate;
            string strEmployee;
            string strManager;
            string strVehicleNumber;
            string strCustomerProjectID;
            string strAssignedProjectID;
            string strProjectName;
            string strOuttime;
            string strInETA;
            string strAssignedOffice;
            string strBuildingAccess;
            string strEmailAddress = "ERPProjectReports@bluejaycommunications.com";
            string strHeader;
            string strBody;

            try
            {
                LoadAfterHoursDataSet();

                strHeader = "After Hours Summary Report For The Last 30 Days";
                strBody = "<h1>" + strHeader + "</h1>";

                strBody += "<table>";
                strBody += "<tr>";
                strBody += "<td><b>Work Date</b></td>";
                strBody += "<td><b>Employee</b></td>";
                strBody += "<td><b>Manager</b></td>";
                strBody += "<td><b>Vehicle Number</b></td>";
                strBody += "<td><b>Customer Project ID</b></td>";
                strBody += "<td><b>Assigned Project ID</b></td>";
                strBody += "<td><b>Project Name</b></td>";
                strBody += "<td><b>Out Time</b></td>";
                strBody += "<td><b>In ETA</b></td>";
                strBody += "<td><b>Assigned Office</b></td>";
                strBody += "<td><b>Building Access</b></td>";
                strBody += "</tr>";

                intNumberOfRecords = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork.Rows.Count;

                if (intNumberOfRecords > 0)
                {
                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strWorkDate = Convert.ToString(TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].WorkDate);
                        strEmployee = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].Employee;
                        strManager = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].Manager;
                        strVehicleNumber = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].VehicleNumber;
                        strCustomerProjectID = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].CustomerProjectID;
                        strAssignedProjectID = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].AssignedProjectID;
                        strProjectName = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].ProjectName;
                        strOuttime = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].OutTime;
                        strInETA = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].InETA;
                        strAssignedOffice = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].AssignedOffice;
                        strBuildingAccess = TheEmployeeAfterHourWorkReportDataSet.employeeafterhourwork[intCounter].BuildingAccess;

                        strBody += "<tr>";
                        strBody += "<td>" + strWorkDate + "</td>";
                        strBody += "<td>" + strEmployee + "</td>";
                        strBody += "<td>" + strManager + "</td>";
                        strBody += "<td>" + strVehicleNumber + "</td>";
                        strBody += "<td>" + strCustomerProjectID + "</td>";
                        strBody += "<td>" + strAssignedProjectID + "</td>";
                        strBody += "<td>" + strProjectName + "</td>";
                        strBody += "<td>" + strOuttime + "</td>";
                        strBody += "<td>" + strInETA + "</td>";
                        strBody += "<td>" + strAssignedOffice + "</td>";
                        strBody += "<td>" + strBuildingAccess + "</td>";
                        strBody += "</tr>";
                    }
                }

                strBody += "</table>";

                TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strBody);

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker //  Run Punched Vs Production // Email Overnight Report " + Ex.ToString());

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

        }
        public void CreateOverTimeReport()
        {
            //setting local variables
            int intCounter;
            int intNumberOfRecords;
            int intManagerID;
            string strManager;
            string strEmailAddress = "ERPProjectReports@bluejaycommunications.com";
            string strHeader;
            string strMessage;
            DateTime datPayDate = DateTime.Now;
            bool blnFatalError = false;

            try
            {
                TheEmployeeOvertimeDataSet.employeeovertime.Rows.Clear();

                datPayDate = TheDateSearchClass.RemoveTime(datPayDate);
                datPayDate = TheDateSearchClass.SubtractingDays(datPayDate, 9);

                TheFindEmployeesOverFortyHoursDataSet = TheEmployeePunchedHoursClass.FindEmployeesOverFortyHours(datPayDate);

                intNumberOfRecords = TheFindEmployeesOverFortyHoursDataSet.FindEmployeesOverFortyHours.Rows.Count;

                if(intNumberOfRecords > 0)
                {
                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        intManagerID = TheFindEmployeesOverFortyHoursDataSet.FindEmployeesOverFortyHours[intCounter].ManagerID;

                        TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(intManagerID);
                        strManager = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].FirstName + " ";
                        strManager += TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].LastName;

                        EmployeeOvertimeDataSet.employeeovertimeRow NewEmployeeRow = TheEmployeeOvertimeDataSet.employeeovertime.NewemployeeovertimeRow();

                        NewEmployeeRow.FirstName = TheFindEmployeesOverFortyHoursDataSet.FindEmployeesOverFortyHours[intCounter].FirstName;
                        NewEmployeeRow.LastName = TheFindEmployeesOverFortyHoursDataSet.FindEmployeesOverFortyHours[intCounter].LastName;
                        NewEmployeeRow.HomeOffice = TheFindEmployeesOverFortyHoursDataSet.FindEmployeesOverFortyHours[intCounter].HomeOffice;
                        NewEmployeeRow.Manager = strManager;
                        NewEmployeeRow.TotalHours = TheFindEmployeesOverFortyHoursDataSet.FindEmployeesOverFortyHours[intCounter].PunchedHours;

                        TheEmployeeOvertimeDataSet.employeeovertime.Rows.Add(NewEmployeeRow);
                    }
                }

                strHeader = "Employee Overtime Report For " + Convert.ToString(datPayDate);

                strMessage = "<h1>Employee Overtime Report For " + Convert.ToString(datPayDate) + "</h1>";
                strMessage += "<p>               </p>";
                strMessage += "<p>               </p>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Home Office</b></td>";
                strMessage += "<td><b>Manager</b></td>";
                strMessage += "<td><b>Total Hours</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                intNumberOfRecords = TheEmployeeOvertimeDataSet.employeeovertime.Rows.Count;

                if (intNumberOfRecords > 0)
                {
                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strMessage += "<tr>";
                        strMessage += "<td>" + TheEmployeeOvertimeDataSet.employeeovertime[intCounter].FirstName + "</td>";
                        strMessage += "<td>" + TheEmployeeOvertimeDataSet.employeeovertime[intCounter].LastName + "</td>";
                        strMessage += "<td>" + TheEmployeeOvertimeDataSet.employeeovertime[intCounter].HomeOffice + "</td>";
                        strMessage += "<td>" + TheEmployeeOvertimeDataSet.employeeovertime[intCounter].Manager + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheEmployeeOvertimeDataSet.employeeovertime[intCounter].TotalHours) + "</td>";
                        strMessage += "</tr>";
                        strMessage += "<p>               </p>";
                    }
                }

                strMessage += "</table>";

                blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Run Punched Vs Production Class // Create Overtime Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }


        }
        public void CreateEmployeeMiscReport()
        {
            //setting local variables
            int intCounter;
            int intNumberOfRecords;
            decimal decTotalHours = 0;
            string strEmailAddress = "ERPProjectReports@bluejaycommunications.com";
            string strHeader;
            string strMessage;
            DateTime datPayDate = DateTime.Now;
            bool blnFatalError = false;
            int intRecordsReturned;

            try
            {
                TheFindEmployeeProductionForMiscCodeDataSet = TheEmployeeProjectAssignmentClass.FindEmployeeProductionForMiscCode();

                intNumberOfRecords = TheFindEmployeesOverFortyHoursDataSet.FindEmployeesOverFortyHours.Rows.Count;

                TheFindEmployeeProductivityMiscTotalHoursDataSet = TheEmployeeProjectAssignmentClass.FindEmplolyeeProductivityMiscTotalHours();

                intRecordsReturned = TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode.Rows.Count;

                if(intRecordsReturned > 0)
                {
                    decTotalHours = TheFindEmployeeProductivityMiscTotalHoursDataSet.FindEmployeeProductivityMiscTotalHours[0].TotalHours;
                }

                strHeader = "Employee Production for Misc with a Total Hours of " + Convert.ToString(decTotalHours);

                strMessage = "<h1>Employee Production for Misc with a Total Hours of " + Convert.ToString(decTotalHours) + "</h1>";
                strMessage += "<p>               </p>";
                strMessage += "<p>               </p>";
                strMessage += "<table>";
                strMessage += "<tr>";
                strMessage += "<td><b>Date</b></td>";
                strMessage += "<td><b>BJC Project ID</b></td>";
                strMessage += "<td><b>Customer Project ID</b></td>";
                strMessage += "<td><b>Project Name</b></td>";
                strMessage += "<td><b>First Name</b></td>";
                strMessage += "<td><b>Last Name</b></td>";
                strMessage += "<td><b>Home Office</b></td>";
                strMessage += "<td><b>Total Hours</b></td>";
                strMessage += "</tr>";
                strMessage += "<p>               </p>";

                if (intNumberOfRecords > 0)
                {
                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strMessage += "<tr>";
                        strMessage += Convert.ToString(TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode[intCounter].TransactionDate);
                        strMessage += "<td>" + TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode[intCounter].AssignedProjectID + "</td>";
                        strMessage += "<td>" + TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode[intCounter].CustomerAssignedID + "</td>";
                        strMessage += "<td>" + TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode[intCounter].ProjectName + "</td>";
                        strMessage += "<td>" + TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode[intCounter].FirstName + "</td>";
                        strMessage += "<td>" + TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode[intCounter].LastName + "</td>";
                        strMessage += "<td>" + TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode[intCounter].HomeOffice + "</td>";
                        strMessage += "<td>" + Convert.ToString(TheFindEmployeeProductionForMiscCodeDataSet.FindEmployeeProductionForMiscCode[intCounter].TotalHours) + "</td>";
                        strMessage += "</tr>";
                        strMessage += "<p>               </p>";
                    }
                }

                strMessage += "</table>";

                blnFatalError = TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);

                if (blnFatalError == true)
                    throw new Exception();

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Run Punched Vs Production Class // Create Employee Misc Report " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
