namespace OutlookAddIn_meetingRoomInfo
{
    partial class MeetingRoomRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Microsoft.Office.Tools.Ribbon.RibbonTab tabMeetingRoom;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup grpQuery;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnQueryToday;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnQueryTomorrow;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnQueryRange;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup grpBooking;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnQuickBook;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabMeetingRoom = this.Factory.CreateRibbonTab();
            this.grpQuery = this.Factory.CreateRibbonGroup();
            this.grpBooking = this.Factory.CreateRibbonGroup();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnQueryToday = this.Factory.CreateRibbonButton();
            this.btnQueryTomorrow = this.Factory.CreateRibbonButton();
            this.btnQueryRange = this.Factory.CreateRibbonButton();
            this.btnQuickBook = this.Factory.CreateRibbonButton();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.tabMeetingRoom.SuspendLayout();
            this.grpQuery.SuspendLayout();
            this.grpBooking.SuspendLayout();
            this.grpHelp.SuspendLayout();
            //
            // tabMeetingRoom
            //
            this.tabMeetingRoom.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Custom;
            this.tabMeetingRoom.Groups.Add(this.grpQuery);
            this.tabMeetingRoom.Groups.Add(this.grpBooking);
            this.tabMeetingRoom.Groups.Add(this.grpHelp);
            this.tabMeetingRoom.Label = "會議室管理";
            this.tabMeetingRoom.Name = "tabMeetingRoom";
            //
            // grpQuery
            //
            this.grpQuery.Items.Add(this.btnQueryToday);
            this.grpQuery.Items.Add(this.btnQueryTomorrow);
            this.grpQuery.Items.Add(this.btnQueryRange);
            this.grpQuery.Label = "查詢功能";
            this.grpQuery.Name = "grpQuery";
            //
            // grpBooking
            //
            this.grpBooking.Items.Add(this.btnQuickBook);
            this.grpBooking.Label = "預約功能";
            this.grpBooking.Name = "grpBooking";
            //
            // grpHelp
            //
            this.grpHelp.Items.Add(this.btnHelp);
            this.grpHelp.Label = "說明";
            this.grpHelp.Name = "grpHelp";
            //
            // btnQueryToday
            //
            this.btnQueryToday.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnQueryToday.Label = "查詢今日";
            this.btnQueryToday.Name = "btnQueryToday";
            this.btnQueryToday.ShowImage = true;
            this.btnQueryToday.SuperTip = "查看今天所有會議室的預約狀況";
            this.btnQueryToday.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQueryToday_Click);
            //
            // btnQueryTomorrow
            //
            this.btnQueryTomorrow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnQueryTomorrow.Label = "查詢明日";
            this.btnQueryTomorrow.Name = "btnQueryTomorrow";
            this.btnQueryTomorrow.ShowImage = true;
            this.btnQueryTomorrow.SuperTip = "查看明天所有會議室的預約狀況";
            this.btnQueryTomorrow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQueryTomorrow_Click);
            //
            // btnQueryRange
            //
            this.btnQueryRange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnQueryRange.Label = "選擇日期";
            this.btnQueryRange.Name = "btnQueryRange";
            this.btnQueryRange.ShowImage = true;
            this.btnQueryRange.SuperTip = "自訂日期範圍查詢預約記錄";
            this.btnQueryRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQueryRange_Click);
            //
            // btnQuickBook
            //
            this.btnQuickBook.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnQuickBook.Label = "快速預約";
            this.btnQuickBook.Name = "btnQuickBook";
            this.btnQuickBook.ShowImage = true;
            this.btnQuickBook.SuperTip = "直接開啟新會議視窗進行預約";
            this.btnQuickBook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQuickBook_Click);
            //
            // btnHelp
            //
            this.btnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHelp.Label = "使用說明";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.ShowImage = true;
            this.btnHelp.SuperTip = "顯示會議室管理系統使用說明";
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelp_Click);
            //
            // MeetingRoomRibbon
            //
            this.Name = "MeetingRoomRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Appointment";
            this.Tabs.Add(this.tabMeetingRoom);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MeetingRoomRibbon_Load);
            this.tabMeetingRoom.ResumeLayout(false);
            this.tabMeetingRoom.PerformLayout();
            this.grpQuery.ResumeLayout(false);
            this.grpQuery.PerformLayout();
            this.grpBooking.ResumeLayout(false);
            this.grpBooking.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();

        }

        #endregion
    }
}
