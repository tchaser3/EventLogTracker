/* Title:           Send Email Class
 * Date:            7-12-17
 * Author:          Terry Holmes */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using NewEventLogDLL;
using System.Configuration;
using System.Net;
using System.Threading;
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
using System.IO;
using SafetyDLL;
using AuditReportEmailListDLL;


namespace EventLogTracker
{
    class SendEmailClass
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        SafetyClass TheSafetyClass = new SafetyClass();
        AuditReportEmailListClass TheAuditReportEmailListClass = new AuditReportEmailListClass();

        VehicleInspectionEmailListDataSet TheVehicleInspectionEmailListDataSet;
        FindSortedAuditEmailListDataSet TheFindSortedAuditEmailListDataSet = new FindSortedAuditEmailListDataSet();

        public bool EmailMessage(string strVehicleNumber, string strVehicleProblem)
        {
            bool blnFatalError = false;
            string strVehicleFailed;
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress;

            try
            {
                strVehicleFailed = strVehicleNumber + " Has Failed The Vehicle Inspection and Can Not Be Driven " + strVehicleProblem;

                TheVehicleInspectionEmailListDataSet = TheSafetyClass.GetVehicleInspectionEmailListInfo();

                intNumberOfRecords = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strEmailAddress = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist[intCounter].EmailAddress;

                    blnFatalError = SendEmail(strEmailAddress, "Vehicle Inspection Failure - Do Not Reply", strVehicleFailed);

                    if (blnFatalError == true)
                        throw new Exception();
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Send Email Class // Email Message " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;
        }
                 
        public bool SendEmail(string mailTo, string subject, string message)
        {
            bool blnFatalError = false;

            try
            {

                MailMessage mailMessage = new MailMessage("bluejayerpreports@bluejaycommunications.com", mailTo, subject, message);
                mailMessage.IsBodyHtml = true;
                mailMessage.BodyEncoding = Encoding.UTF8;
                mailMessage.SubjectEncoding = Encoding.UTF8;

                SmtpClient smtpClient = new SmtpClient("192.168.0.240", 25);
                smtpClient.UseDefaultCredentials = false;
                smtpClient.EnableSsl = false;
                smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtpClient.Send(mailMessage);
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Send Email Class // Send Mail " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;

        }
        public bool VehicleReports(string strHeader, String strVehicleReport)
        {
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress;
            bool blnFatalError = false;

            try
            {
                TheVehicleInspectionEmailListDataSet = TheSafetyClass.GetVehicleInspectionEmailListInfo();

                intNumberOfRecords = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strEmailAddress = TheVehicleInspectionEmailListDataSet.vehicleinspectionemaillist[intCounter].EmailAddress;

                    blnFatalError = SendEmail(strEmailAddress, strHeader, strVehicleReport);

                    if (blnFatalError == true)
                        throw new Exception();
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Event Log Tracker // Send Email Class // Vehicle Reports " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());

                blnFatalError = true;
            }

            return blnFatalError;            
        }
       
    }
}
