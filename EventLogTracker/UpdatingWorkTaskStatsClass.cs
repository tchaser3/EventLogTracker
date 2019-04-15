/* Title:           Update Work Task Stats Class
 * Date:            3-26-18
 * Author:          Terry Holmes
 * 
 * Description:     This class will update the Work Task Stats Table */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NewEventLogDLL;
using DateSearchDLL;
using EmployeeProjectAssignmentDLL;
using WorkTaskDLL;
using WorkTaskStatsDLL;

namespace EventLogTracker
{
    class UpdatingWorkTaskStatsClass
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        EmployeeProjectAssignmentClass TheEmployeeProjectAssignmentClass = new EmployeeProjectAssignmentClass();
        WorkTaskClass TheWorkTaskClass = new WorkTaskClass();
        WorkTaskStatsClass TheWorkTaskStatsClass = new WorkTaskStatsClass();

        FindLaborHoursByDateRangeDataSet TheFindLaborHoursByDateRangeDataSet = new FindLaborHoursByDateRangeDataSet();
        FindWorkTaskHoursDataSet TheFindWorkTaskHoursDataSet = new FindWorkTaskHoursDataSet();
        FindWorkTaskStatsByTaskIDDataSet TheFindWorkTaskStatsByTaskIDDataSet = new FindWorkTaskStatsByTaskIDDataSet();
        WorkStatsDataSet TheWorkStatsDataSet = new WorkStatsDataSet();

        decimal gdecTotal;
        decimal gdecMean;
        int gintCounter;
        decimal gdecVariance;
        decimal gdecStandardDeviation;
        decimal gdecLimiter;
        int gintTaskUpperLimit;

        public bool UpdateWorkTaskStatsTable()
        {
            bool blnFatalError = false;
            DateTime datStartDate = DateTime.Now;
            DateTime datEndDate = DateTime.Now;
            int intCounter;
            int intNumberOfRecords;
            double douHours;
            int intTaskCounter;
            bool blnItemFound;
            int intWorkTaskID;
            decimal decMean;
            decimal decVariance;
            decimal dectotalHours;
            int intItems;
            string strWorkTask;
            double douTaskHours;
            decimal decStandardDeviation;
            int intRecordsReturned;
            decimal decLimiter;

            try
            {
                TheWorkStatsDataSet.workstats.Rows.Clear();

                datEndDate = TheDateSearchClass.RemoveTime(datEndDate);

                datStartDate = TheDateSearchClass.SubtractingDays(datEndDate, 900);

                gintCounter = 0;
                gdecTotal = 0;
                gdecVariance = 0;
                gintTaskUpperLimit = 0;

                TheFindLaborHoursByDateRangeDataSet = TheEmployeeProjectAssignmentClass.FindLaborHoursByDateRange(datStartDate, datEndDate);

                intNumberOfRecords = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    if (TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours > 0)
                    {
                        intWorkTaskID = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].WorkTaskID;
                        blnItemFound = false;

                        if (gintTaskUpperLimit > 0)
                        {
                            for (intTaskCounter = 0; intTaskCounter < gintTaskUpperLimit; intTaskCounter++)
                            {
                                if (intWorkTaskID == TheWorkStatsDataSet.workstats[intTaskCounter].WorkTaskID)
                                {
                                    blnItemFound = true;
                                    TheWorkStatsDataSet.workstats[intTaskCounter].HoursPerTask += TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours;
                                    TheWorkStatsDataSet.workstats[intTaskCounter].ItemCounter++;
                                }
                            }
                        }

                        gdecTotal += TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours;
                        gintCounter++;

                        if (blnItemFound == false)
                        {
                            WorkStatsDataSet.workstatsRow NewTaskRow = TheWorkStatsDataSet.workstats.NewworkstatsRow();

                            NewTaskRow.HoursPerTask = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours;
                            NewTaskRow.ItemCounter = 1;
                            NewTaskRow.Limiter = 0;
                            NewTaskRow.Mean = 0;
                            NewTaskRow.StandardDeviation = 0;
                            NewTaskRow.Variance = 0;
                            NewTaskRow.WorkTask = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].WorkTask;
                            NewTaskRow.WorkTaskID = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].WorkTaskID;

                            TheWorkStatsDataSet.workstats.Rows.Add(NewTaskRow);
                            gintTaskUpperLimit++;
                        }
                    }
                }

                decVariance = 0;

                for (intTaskCounter = 0; intTaskCounter < gintTaskUpperLimit; intTaskCounter++)
                {
                    intItems = TheWorkStatsDataSet.workstats[intTaskCounter].ItemCounter;
                    dectotalHours = TheWorkStatsDataSet.workstats[intTaskCounter].HoursPerTask;
                    intWorkTaskID = TheWorkStatsDataSet.workstats[intTaskCounter].WorkTaskID;
                    strWorkTask = TheWorkStatsDataSet.workstats[intTaskCounter].WorkTask;

                    decMean = dectotalHours / intItems;

                    decMean = Math.Round(decMean, 4);

                    TheWorkStatsDataSet.workstats[intTaskCounter].Mean = decMean;

                    TheFindWorkTaskHoursDataSet = TheWorkTaskClass.FindWorkTaskHours(strWorkTask, datStartDate, datEndDate);

                    intNumberOfRecords = TheFindWorkTaskHoursDataSet.FindWorkTaskHours.Rows.Count - 1;

                    for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        if (TheFindWorkTaskHoursDataSet.FindWorkTaskHours[intCounter].TotalHours > 0)
                        {
                            douTaskHours = Convert.ToDouble(TheFindWorkTaskHoursDataSet.FindWorkTaskHours[intCounter].TotalHours - decMean);

                            decVariance += Convert.ToDecimal(Math.Pow(douTaskHours, 2));
                        }
                    }

                    decVariance = decVariance / intItems;

                    decVariance = Math.Round(decVariance, 4);

                    decStandardDeviation = Math.Round(Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(decVariance))), 4);

                    TheWorkStatsDataSet.workstats[intTaskCounter].Variance = decVariance;

                    TheWorkStatsDataSet.workstats[intTaskCounter].StandardDeviation = decStandardDeviation;

                    TheWorkStatsDataSet.workstats[intTaskCounter].Limiter = decMean + (5 * decStandardDeviation);
                }

                gdecMean = gdecTotal / gintCounter;

                gdecMean = Math.Round(gdecMean, 4);

                intNumberOfRecords = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    if (TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours > 0)
                    {
                        douHours = Convert.ToDouble(TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours - gdecMean);

                        gdecVariance += Convert.ToDecimal(Math.Pow(douHours, 2));
                    }
                }

                gdecVariance = gdecVariance / gintCounter;

                gdecStandardDeviation = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(gdecVariance)));

                gdecStandardDeviation = Math.Round(gdecStandardDeviation, 4);

                gdecLimiter = gdecMean + (gdecStandardDeviation * 5);

                //adding to the table
                intNumberOfRecords = TheWorkStatsDataSet.workstats.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intWorkTaskID = TheWorkStatsDataSet.workstats[intCounter].WorkTaskID;
                    decMean = TheWorkStatsDataSet.workstats[intCounter].Mean;
                    decVariance = TheWorkStatsDataSet.workstats[intCounter].Variance;
                    decStandardDeviation = TheWorkStatsDataSet.workstats[intCounter].StandardDeviation;
                    decLimiter = TheWorkStatsDataSet.workstats[intCounter].Limiter;

                    TheFindWorkTaskStatsByTaskIDDataSet = TheWorkTaskStatsClass.FindWorkTaskStatsByTaskID(intWorkTaskID);

                    intRecordsReturned = TheFindWorkTaskStatsByTaskIDDataSet.FindWorkTaskStatsByWorkTaskID.Rows.Count;

                    if(intRecordsReturned > 0)
                    {
                        blnFatalError = TheWorkTaskStatsClass.UpdatetWorkTaskStats(intWorkTaskID, decMean, decVariance, decStandardDeviation, decLimiter);

                        if (blnFatalError == true)
                            throw new Exception();
                    }
                    if(intRecordsReturned == 0)
                    {
                        blnFatalError = TheWorkTaskStatsClass.InsertWorkTaskStats(intWorkTaskID, decMean, decVariance, decStandardDeviation, decLimiter);

                        if (blnFatalError == true)
                            throw new Exception();
                    }
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Update Work Task Stats Class // Update Work Task Stats Table " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
    }
}
