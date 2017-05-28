using System.Runtime.InteropServices;
using Hardcodet.Wpf.TaskbarNotification;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailNotifierAddin
{
    public partial class ThisAddIn 
    {
        private const int BalloonTimeout = 5000;
        private Outlook.Inspectors _inspectors;
        private TaskbarIcon _taskbarIcon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _taskbarIcon = new TaskbarIcon
            {
                Icon = Icons.Zerode_Plump_Mail,
            };
            _inspectors = Application.Inspectors;
            //_inspectors.NewInspector += Inspectors_NewInspector;
            Application.NewMailEx += NewMailExEvent;

        }

        private void NewMailExEvent(string entryIdCollection)
        {
            try
            {

                var mail = Application.Session.GetItemFromID(entryIdCollection) as Outlook.MailItem;
                if (mail != null)
                {
                    // A mail
                    _taskbarIcon.TrayBalloonTipClicked += (sender, args) =>
                    {
                        mail.Display();
                    };

                    _taskbarIcon.ShowBalloonTip(mail.Subject, $"From: {mail.SenderName}", BalloonIcon.Info);
                    return;
                }

                var meeting = Application.Session.GetItemFromID(entryIdCollection) as Outlook.MeetingItem;
                if (meeting != null)
                {
                    // A meeting
                    _taskbarIcon.TrayBalloonTipClicked += (sender, args) =>
                    {
                        meeting.Display();
                    };

                    _taskbarIcon.ShowBalloonTip(meeting.Subject, $"From: {meeting.SenderName}", BalloonIcon.Info);
                    return;
                }
            }
            catch (COMException e)
            {
                // Ignore
            }

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
            _taskbarIcon.Dispose();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
