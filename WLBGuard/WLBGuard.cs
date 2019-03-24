using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace WLBGuard
{
    public partial class WLBGuard
    {
        private static TimeSpan s_minSendTime = TimeSpan.FromHours(8);
        private static TimeSpan s_maxSendTime = TimeSpan.FromHours(20);
        private static DayOfWeek s_firstWorkDay = DayOfWeek.Sunday;
        private static DayOfWeek s_lastWorkDay = DayOfWeek.Thursday;

        private static DateTime s_outlookMagicDateTimeNotDefinedConst = new DateTime(4501, 1, 1, 0, 0, 0);

        private void WLBGuard_Startup(object sender, EventArgs e)
        {
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            var item = new OutlookItem(Item);

            if (!item.IsSupportedItem)
            {
                // non-supported item, for now ...
                return;
            }

            if (item.DeferredDeliveryTime != s_outlookMagicDateTimeNotDefinedConst)
            {
                // already deferred, don't get involve
                return;
            }

            DateTime deferredDeliveryTime = GetNextAllowedTime();
          
            if (deferredDeliveryTime <= DateTime.Now)
            {
                // we are in the allowed time
                return;
            }

            DialogResult result = MessageBox.Show($"It's off work time, defer this item to {deferredDeliveryTime}?", "Work-life Balance Guard", MessageBoxButtons.YesNoCancel);//, , button, icon);
            
            switch (result)
            {
                case DialogResult.Yes:
                    item.DeferredDeliveryTime = deferredDeliveryTime;
                    break;
                case DialogResult.No:
                    // do nothing
                    break;
                case DialogResult.Cancel:
                    Cancel = true; 
                    break;
            }            
        }

        private DateTime GetNextAllowedTime()
        {
            var now = DateTime.Now;
            
            if (AllowedToSend(now))
            {
                // Kosher mail
                return now;
            }

            DateTime nextAllowed = now.Date.AddDays(now.TimeOfDay > s_maxSendTime ? 1 : 0) + s_minSendTime;

            while (!AllowedToSend(nextAllowed))
            {
                nextAllowed = nextAllowed.AddDays(1);
            }

            return nextAllowed;
        }

        private static bool AllowedToSend(DateTime time)
        {
            TimeSpan timeOfDay = time.TimeOfDay;
            DayOfWeek dayOfWeek = time.DayOfWeek;

            return timeOfDay >= s_minSendTime && timeOfDay <= s_maxSendTime && dayOfWeek >= s_firstWorkDay && dayOfWeek <= s_lastWorkDay;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new EventHandler(WLBGuard_Startup);
        }

        #endregion


        /// <summary>
        /// Outlook add-in is using COM, in order to simplify calling code this wraps commonan properties of different wrapped items
        /// e.g. Mail and Meeting DeferredDeliveryTime
        /// </summary>
        private class OutlookItem
        {
            private readonly Outlook.MailItem m_mail;
            private readonly Outlook.MeetingItem m_meeting;

            public OutlookItem(object Item)
            {
                m_mail = Item as Outlook.MailItem;
                m_meeting = Item as Outlook.MeetingItem;

                IsSupportedItem = m_mail != null || m_meeting != null;
            }

            // Any call to properties when IsSupportedItem is false is unexpected
            public bool IsSupportedItem { get; private set; }

            public DateTime DeferredDeliveryTime
            {
                get
                {
                    return m_mail != null ? m_mail.DeferredDeliveryTime : m_meeting.DeferredDeliveryTime;
                }
                set
                {
                    if (m_mail != null)
                    {
                        m_mail.DeferredDeliveryTime = value;
                    }
                    else
                    {
                        m_meeting.DeferredDeliveryTime = value;
                    }
                }
            }
        }
    }
}
