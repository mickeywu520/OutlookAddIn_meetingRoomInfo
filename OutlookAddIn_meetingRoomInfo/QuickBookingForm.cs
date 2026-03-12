using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace OutlookAddIn_meetingRoomInfo
{
    public partial class QuickBookingForm : Form
    {
        private List<MeetingRecord> _allRecords;
        private List<MeetingRoom> _rooms;
        private DateTime _selectedDate;
        private string _selectedRoomId;
        private DateTime _selectedStartTime;
        private DateTime _selectedEndTime;

        // 用於重新載入資料的委派
        private Func<DateTime, DateTime, Task<List<MeetingRecord>>> _fetchRecordsFunc;

        // Output properties
        public string SelectedRoomId => _selectedRoomId;
        public string SelectedRoomDisplayName { get; private set; }
        public DateTime SelectedStartTime => _selectedStartTime;
        public DateTime SelectedEndTime => _selectedEndTime;
        public string MeetingSubject { get; private set; }

        // 週期性預約相關屬性
        public bool IsRecurrentBooking { get; private set; }
        public RecurrenceSettings RecurrenceSettings { get; private set; }
        public BatchBookingResult BatchBookingResult { get; private set; }

        // 用於立即預約會議室的委派
        private Func<string, string, DateTime, DateTime, string, Task<bool>> _bookRoomFunc;

        // TabControl
        private TabControl tabControl;
        private TabPage tabSingle;
        private TabPage tabRecurrent;

        // 單次預約控制項
        private ComboBox cmbRooms;
        private DateTimePicker dtpDate;
        private DataGridView dgvAvailableSlots;
        private Button btnBook;
        private Button btnCancel;
        private Label lblRoom;
        private Label lblDate;
        private Label lblTitle;
        private Label lblRemark;

        // 週期預約控制項
        private ComboBox cmbRecRooms;
        private DateTimePicker dtpRecStartDate;
        private ComboBox cmbRecurrenceType;
        private NumericUpDown numInterval;
        private CheckedListBox clbDaysOfWeek;
        private RadioButton rdoEndByDate;
        private RadioButton rdoEndByOccurrences;
        private DateTimePicker dtpEndDate;
        private NumericUpDown numOccurrences;
        private ComboBox cmbTimeSlot;
        private DataGridView dgvRecPreview;
        private Button btnGeneratePreview;
        private Button btnBookRecurrent;
        private Button btnCancelRecurrent;
        private Label lblRecRoom;
        private Label lblRecStartDate;
        private Label lblRecType;
        private Label lblRecInterval;
        private Label lblRecDays;
        private Label lblRecEnd;
        private Label lblRecTimeSlot;
        private Label lblRecPreview;

        public QuickBookingForm(List<MeetingRecord> existingRecords, List<MeetingRoom> rooms,
            Func<DateTime, DateTime, Task<List<MeetingRecord>>> fetchRecordsFunc = null,
            Func<string, string, DateTime, DateTime, string, Task<bool>> bookRoomFunc = null)
        {
            _allRecords = existingRecords ?? new List<MeetingRecord>();
            _rooms = rooms ?? new List<MeetingRoom>();
            _fetchRecordsFunc = fetchRecordsFunc;
            _bookRoomFunc = bookRoomFunc;
            _selectedDate = DateTime.Now;
            IsRecurrentBooking = false;
            InitializeComponent();
            LoadRooms();
            LoadRecRooms();
            RefreshAvailableSlots();
        }

        private void InitializeComponent()
        {
            this.Text = "快速預約會議室";
            this.Size = new Size(900, 700);
            this.MinimumSize = new Size(850, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;

            // Title
            lblTitle = new Label();
            lblTitle.Text = "選擇會議室與時段";
            lblTitle.Font = new Font("Microsoft JhengHei", 14, FontStyle.Bold);
            lblTitle.Location = new Point(20, 15);
            lblTitle.Size = new Size(300, 30);
            this.Controls.Add(lblTitle);

            // TabControl
            tabControl = new TabControl();
            tabControl.Location = new Point(20, 50);
            tabControl.Size = new Size(850, 580);
            tabControl.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(tabControl);

            // 初始化單次預約頁籤
            InitializeSingleBookingTab();

            // 初始化週期預約頁籤
            InitializeRecurrentBookingTab();

            // 設定表單層級的 Cancel/Accept 按鈕
            this.CancelButton = btnCancel;
            this.AcceptButton = btnBook;
        }

        private void InitializeSingleBookingTab()
        {
            tabSingle = new TabPage("單次預約");
            tabSingle.Padding = new Padding(0);
            tabControl.TabPages.Add(tabSingle);

            // =============================================
            // 底部按鈕 Panel（DockStyle.Bottom，固定高度）
            // 必須先加入，才能讓 Fill Panel 正確填滿剩餘空間
            // =============================================
            Panel bottomPanel = new Panel();
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.Height = 50;
            bottomPanel.Padding = new Padding(5);
            tabSingle.Controls.Add(bottomPanel);

            // Cancel button
            btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Size = new Size(100, 35);
            btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            bottomPanel.Controls.Add(btnCancel);

            // Book button
            btnBook = new Button();
            btnBook.Text = "預約";
            btnBook.Size = new Size(100, 35);
            btnBook.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnBook.Enabled = false;
            btnBook.Click += BtnBook_Click;
            bottomPanel.Controls.Add(btnBook);

            // 動態調整按鈕位置（靠右對齊）
            bottomPanel.Resize += (s, e) =>
            {
                btnCancel.Left = bottomPanel.Width - 110;
                btnCancel.Top  = 7;
                btnBook.Left   = bottomPanel.Width - 220;
                btnBook.Top    = 7;
            };
            // 初始定位（Resize 事件在第一次顯示前不會觸發）
            bottomPanel.Width = tabSingle.Width;
            btnCancel.Left = bottomPanel.Width - 110;
            btnCancel.Top  = 7;
            btnBook.Left   = bottomPanel.Width - 220;
            btnBook.Top    = 7;

            // =============================================
            // 內容 Panel（DockStyle.Fill，填滿剩餘空間）
            // =============================================
            Panel contentPanel = new Panel();
            contentPanel.Dock = DockStyle.Fill;
            contentPanel.Padding = new Padding(10);
            tabSingle.Controls.Add(contentPanel);

            // Room label
            lblRoom = new Label();
            lblRoom.Text = "會議室:";
            lblRoom.Location = new Point(20, 20);
            lblRoom.Size = new Size(80, 25);
            contentPanel.Controls.Add(lblRoom);

            // Room combo box
            cmbRooms = new ComboBox();
            cmbRooms.Location = new Point(110, 18);
            cmbRooms.Size = new Size(280, 25);
            cmbRooms.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbRooms.SelectedIndexChanged += CmbRooms_SelectedIndexChanged;
            contentPanel.Controls.Add(cmbRooms);

            // Remark label (to show room info)
            lblRemark = new Label();
            lblRemark.Location = new Point(400, 20);
            lblRemark.Size = new Size(400, 25);
            lblRemark.Font = new Font("Microsoft JhengHei", 9, FontStyle.Italic);
            lblRemark.ForeColor = Color.Gray;
            contentPanel.Controls.Add(lblRemark);

            // Date label
            lblDate = new Label();
            lblDate.Text = "日期:";
            lblDate.Location = new Point(20, 55);
            lblDate.Size = new Size(50, 25);
            contentPanel.Controls.Add(lblDate);

            // Date picker
            dtpDate = new DateTimePicker();
            dtpDate.Location = new Point(110, 53);
            dtpDate.Size = new Size(150, 25);
            dtpDate.Format = DateTimePickerFormat.Short;
            dtpDate.MinDate = DateTime.Now.Date;
            dtpDate.ValueChanged += DtpDate_ValueChanged;
            contentPanel.Controls.Add(dtpDate);

            // Available slots grid（填滿 contentPanel 剩餘空間）
            dgvAvailableSlots = new DataGridView();
            dgvAvailableSlots.Location = new Point(20, 90);
            dgvAvailableSlots.Anchor = AnchorStyles.Top | AnchorStyles.Bottom
                                     | AnchorStyles.Left | AnchorStyles.Right;
            /*
            dgvAvailableSlots.Size = new Size(
                contentPanel.Width  - 40,
                contentPanel.Height - 100);
            */
            // 修改後 — 給一個合理的固定初始值，表單顯示後 Anchor 自動撐開
            dgvAvailableSlots.Size = new Size(800, 430);  
            dgvAvailableSlots.AllowUserToAddRows    = false;
            dgvAvailableSlots.AllowUserToDeleteRows = false;
            dgvAvailableSlots.ReadOnly              = true;
            dgvAvailableSlots.SelectionMode         = DataGridViewSelectionMode.FullRowSelect;
            dgvAvailableSlots.AutoSizeColumnsMode   = DataGridViewAutoSizeColumnsMode.Fill;
            dgvAvailableSlots.RowHeadersVisible     = false;
            dgvAvailableSlots.MultiSelect           = true;
            dgvAvailableSlots.SelectionChanged      += DgvAvailableSlots_SelectionChanged;
            dgvAvailableSlots.CellDoubleClick       += DgvAvailableSlots_CellDoubleClick;
            contentPanel.Controls.Add(dgvAvailableSlots);

            // Setup columns
            dgvAvailableSlots.Columns.Add("TimeSlot", "時段");
            dgvAvailableSlots.Columns.Add("Status",   "狀態");
            dgvAvailableSlots.Columns.Add("Booker",   "預約人");
            dgvAvailableSlots.Columns.Add("Subject",  "會議主題");
            dgvAvailableSlots.Columns.Add("Duration", "時長");
        }

        private void InitializeRecurrentBookingTab()
        {
            tabRecurrent = new TabPage("週期預約");
            tabRecurrent.Padding = new Padding(0);
            tabControl.TabPages.Add(tabRecurrent);

            // =============================================
            // 底部按鈕 Panel（DockStyle.Bottom，固定高度）
            // 必須先加入，才能讓 Fill Panel 正確填滿剩餘空間
            // =============================================
            Panel recBottomPanel = new Panel();
            recBottomPanel.Dock   = DockStyle.Bottom;
            recBottomPanel.Height = 50;
            recBottomPanel.Padding = new Padding(5);
            tabRecurrent.Controls.Add(recBottomPanel);

            // Cancel button
            btnCancelRecurrent = new Button();
            btnCancelRecurrent.Text         = "取消";
            btnCancelRecurrent.DialogResult = DialogResult.Cancel;
            btnCancelRecurrent.Size         = new Size(100, 35);
            btnCancelRecurrent.Anchor       = AnchorStyles.Right | AnchorStyles.Top;
            recBottomPanel.Controls.Add(btnCancelRecurrent);

            // Book recurrent button
            btnBookRecurrent = new Button();
            btnBookRecurrent.Text    = "批次預約";
            btnBookRecurrent.Size    = new Size(100, 35);
            btnBookRecurrent.Anchor  = AnchorStyles.Right | AnchorStyles.Top;
            btnBookRecurrent.Enabled = false;
            btnBookRecurrent.Click  += BtnBookRecurrent_Click;
            recBottomPanel.Controls.Add(btnBookRecurrent);

            // 動態調整按鈕位置（靠右對齊）
            recBottomPanel.Resize += (s, e) =>
            {
                btnCancelRecurrent.Left = recBottomPanel.Width - 110;
                btnCancelRecurrent.Top  = 7;
                btnBookRecurrent.Left   = recBottomPanel.Width - 220;
                btnBookRecurrent.Top    = 7;
            };
            // 初始定位
            recBottomPanel.Width        = tabRecurrent.Width;
            btnCancelRecurrent.Left     = recBottomPanel.Width - 110;
            btnCancelRecurrent.Top      = 7;
            btnBookRecurrent.Left       = recBottomPanel.Width - 220;
            btnBookRecurrent.Top        = 7;

            // =============================================
            // 內容 Panel（DockStyle.Fill，填滿剩餘空間）
            // =============================================
            Panel recContentPanel = new Panel();
            recContentPanel.Dock      = DockStyle.Fill;
            recContentPanel.AutoScroll = true; // 內容較多時允許捲動
            recContentPanel.Padding   = new Padding(10);
            tabRecurrent.Controls.Add(recContentPanel);

            int yPos = 20;

            // Room label and combo
            lblRecRoom = new Label();
            lblRecRoom.Text     = "會議室:";
            lblRecRoom.Location = new Point(20, yPos);
            lblRecRoom.Size     = new Size(80, 25);
            recContentPanel.Controls.Add(lblRecRoom);

            cmbRecRooms = new ComboBox();
            cmbRecRooms.Location      = new Point(110, yPos - 2);
            cmbRecRooms.Size          = new Size(280, 25);
            cmbRecRooms.DropDownStyle = ComboBoxStyle.DropDownList;
            recContentPanel.Controls.Add(cmbRecRooms);

            yPos += 35;

            // Start date
            lblRecStartDate = new Label();
            lblRecStartDate.Text     = "開始日期:";
            lblRecStartDate.Location = new Point(20, yPos);
            lblRecStartDate.Size     = new Size(80, 25);
            recContentPanel.Controls.Add(lblRecStartDate);

            dtpRecStartDate = new DateTimePicker();
            dtpRecStartDate.Location = new Point(110, yPos - 2);
            dtpRecStartDate.Size     = new Size(150, 25);
            dtpRecStartDate.Format   = DateTimePickerFormat.Short;
            dtpRecStartDate.MinDate  = DateTime.Now.Date;
            recContentPanel.Controls.Add(dtpRecStartDate);

            yPos += 35;

            // Recurrence type
            lblRecType = new Label();
            lblRecType.Text     = "重複頻率:";
            lblRecType.Location = new Point(20, yPos);
            lblRecType.Size     = new Size(80, 25);
            recContentPanel.Controls.Add(lblRecType);

            cmbRecurrenceType = new ComboBox();
            cmbRecurrenceType.Location      = new Point(110, yPos - 2);
            cmbRecurrenceType.Size          = new Size(120, 25);
            cmbRecurrenceType.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbRecurrenceType.Items.AddRange(new object[] { "每日", "每週", "每月" });
            cmbRecurrenceType.SelectedIndex         = 1; // 預設每週
            cmbRecurrenceType.SelectedIndexChanged += CmbRecurrenceType_SelectedIndexChanged;
            recContentPanel.Controls.Add(cmbRecurrenceType);

            // Interval prefix label
            lblRecInterval = new Label();
            lblRecInterval.Text     = "每";
            lblRecInterval.Location = new Point(240, yPos);
            lblRecInterval.Size     = new Size(25, 25);
            recContentPanel.Controls.Add(lblRecInterval);

            numInterval = new NumericUpDown();
            numInterval.Location = new Point(270, yPos - 2);
            numInterval.Size     = new Size(50, 25);
            numInterval.Minimum  = 1;
            numInterval.Maximum  = 52;
            numInterval.Value    = 1;
            recContentPanel.Controls.Add(numInterval);

            // Interval unit label
            lblRecInterval = new Label();
            lblRecInterval.Text     = "週";
            lblRecInterval.Location = new Point(325, yPos);
            lblRecInterval.Size     = new Size(30, 25);
            lblRecInterval.Name     = "lblIntervalUnit";
            recContentPanel.Controls.Add(lblRecInterval);

            yPos += 35;

            // Days of week (for weekly)
            lblRecDays = new Label();
            lblRecDays.Text     = "星期:";
            lblRecDays.Location = new Point(20, yPos);
            lblRecDays.Size     = new Size(80, 25);
            recContentPanel.Controls.Add(lblRecDays);

            clbDaysOfWeek = new CheckedListBox();
            clbDaysOfWeek.Location = new Point(110, yPos - 2);
            clbDaysOfWeek.Size     = new Size(280, 80);
            clbDaysOfWeek.Items.AddRange(new object[]
                { "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日" });
            clbDaysOfWeek.CheckOnClick = true;
            clbDaysOfWeek.SetItemChecked(0, true); // 預設勾選星期一
            recContentPanel.Controls.Add(clbDaysOfWeek);

            yPos += 90;

            // End condition
            lblRecEnd = new Label();
            lblRecEnd.Text     = "結束方式:";
            lblRecEnd.Location = new Point(20, yPos);
            lblRecEnd.Size     = new Size(80, 25);
            recContentPanel.Controls.Add(lblRecEnd);

            rdoEndByDate = new RadioButton();
            rdoEndByDate.Text     = "結束日期:";
            rdoEndByDate.Location = new Point(110, yPos - 2);
            rdoEndByDate.Size     = new Size(90, 25);
            rdoEndByDate.Checked  = true;
            recContentPanel.Controls.Add(rdoEndByDate);

            dtpEndDate = new DateTimePicker();
            dtpEndDate.Location = new Point(205, yPos - 2);
            dtpEndDate.Size     = new Size(150, 25);
            dtpEndDate.Format   = DateTimePickerFormat.Short;
            dtpEndDate.MinDate  = DateTime.Now.Date.AddDays(1);
            dtpEndDate.Value    = DateTime.Now.Date.AddMonths(1);
            recContentPanel.Controls.Add(dtpEndDate);

            yPos += 30;

            rdoEndByOccurrences = new RadioButton();
            rdoEndByOccurrences.Text     = "重複次數:";
            rdoEndByOccurrences.Location = new Point(110, yPos - 2);
            rdoEndByOccurrences.Size     = new Size(90, 25);
            recContentPanel.Controls.Add(rdoEndByOccurrences);

            numOccurrences = new NumericUpDown();
            numOccurrences.Location = new Point(205, yPos - 2);
            numOccurrences.Size     = new Size(60, 25);
            numOccurrences.Minimum  = 1;
            numOccurrences.Maximum  = 100;
            numOccurrences.Value    = 10;
            numOccurrences.Enabled  = false;
            recContentPanel.Controls.Add(numOccurrences);

            rdoEndByDate.CheckedChanged += (s, e) =>
            {
                dtpEndDate.Enabled     = rdoEndByDate.Checked;
                numOccurrences.Enabled = rdoEndByOccurrences.Checked;
            };
            rdoEndByOccurrences.CheckedChanged += (s, e) =>
            {
                dtpEndDate.Enabled     = rdoEndByDate.Checked;
                numOccurrences.Enabled = rdoEndByOccurrences.Checked;
            };

            yPos += 35;

            // Time slot
            lblRecTimeSlot = new Label();
            lblRecTimeSlot.Text     = "固定時段:";
            lblRecTimeSlot.Location = new Point(20, yPos);
            lblRecTimeSlot.Size     = new Size(80, 25);
            recContentPanel.Controls.Add(lblRecTimeSlot);

            cmbTimeSlot = new ComboBox();
            cmbTimeSlot.Location      = new Point(110, yPos - 2);
            cmbTimeSlot.Size          = new Size(150, 25);
            cmbTimeSlot.DropDownStyle = ComboBoxStyle.DropDownList;
            for (int hour = 8; hour <= 18; hour++)
            {
                cmbTimeSlot.Items.Add(string.Format("{0:00}:00 - {0:00}:30", hour));
                if (hour < 18)
                    cmbTimeSlot.Items.Add(string.Format("{0:00}:30 - {1:00}:00", hour, hour + 1));
            }
            cmbTimeSlot.SelectedIndex = 2; // 預設 08:30-09:00
            recContentPanel.Controls.Add(cmbTimeSlot);

            // Generate preview button
            btnGeneratePreview = new Button();
            btnGeneratePreview.Text     = "產生預覽";
            btnGeneratePreview.Location = new Point(280, yPos - 2);
            btnGeneratePreview.Size     = new Size(100, 25);
            btnGeneratePreview.Click   += BtnGeneratePreview_Click;
            recContentPanel.Controls.Add(btnGeneratePreview);

            yPos += 35;

            // Preview label
            lblRecPreview = new Label();
            lblRecPreview.Text     = "預覽:";
            lblRecPreview.Location = new Point(20, yPos);
            lblRecPreview.Size     = new Size(80, 25);
            recContentPanel.Controls.Add(lblRecPreview);

            yPos += 25;

            // Preview grid（填滿 recContentPanel 剩餘空間）
            dgvRecPreview = new DataGridView();
            dgvRecPreview.Location = new Point(20, yPos);
            dgvRecPreview.Anchor   = AnchorStyles.Top | AnchorStyles.Bottom
                                   | AnchorStyles.Left | AnchorStyles.Right;
            /*
            dgvRecPreview.Size = new Size(
                recContentPanel.Width  - 40,
                recContentPanel.Height - yPos - 10);
            */
            dgvRecPreview.Size = new Size(800, 150);   // Anchor 自動撐開
            dgvRecPreview.AllowUserToAddRows    = false;
            dgvRecPreview.AllowUserToDeleteRows = false;
            dgvRecPreview.ReadOnly              = true;
            dgvRecPreview.SelectionMode         = DataGridViewSelectionMode.FullRowSelect;
            dgvRecPreview.AutoSizeColumnsMode   = DataGridViewAutoSizeColumnsMode.Fill;
            dgvRecPreview.RowHeadersVisible     = false;
            recContentPanel.Controls.Add(dgvRecPreview);

            // Setup columns
            dgvRecPreview.Columns.Add("Date",   "日期");
            dgvRecPreview.Columns.Add("Time",   "時段");
            dgvRecPreview.Columns.Add("Status", "狀態");
            dgvRecPreview.Columns.Add("Room",   "會議室");
        }

        private void CmbRecurrenceType_SelectedIndexChanged(object sender, EventArgs e)
        {
            var lblUnit = tabRecurrent.Controls
                .OfType<Panel>()
                .SelectMany(p => p.Controls.OfType<Label>())
                .FirstOrDefault(l => l.Name == "lblIntervalUnit");

            if (lblUnit != null)
            {
                switch (cmbRecurrenceType.SelectedIndex)
                {
                    case 0: // 每日
                        lblUnit.Text           = "天";
                        clbDaysOfWeek.Enabled  = false;
                        break;
                    case 1: // 每週
                        lblUnit.Text           = "週";
                        clbDaysOfWeek.Enabled  = true;
                        break;
                    case 2: // 每月
                        lblUnit.Text           = "月";
                        clbDaysOfWeek.Enabled  = false;
                        break;
                }
            }
        }

        private void LoadRooms()
        {
            cmbRooms.Items.Clear();

            foreach (var room in _rooms.Where(r => !r.Disable).OrderBy(r => r.Sort))
            {
                var item = new RoomComboItem
                {
                    RoomId      = room.RoomId,
                    Name        = room.Name,
                    Remark      = room.Remark,
                    DisplayName = string.Format("{0} - {1}", room.RoomId, room.Name)
                };
                cmbRooms.Items.Add(item);
            }

            if (cmbRooms.Items.Count > 0)
                cmbRooms.SelectedIndex = 0;
        }

        private void LoadRecRooms()
        {
            cmbRecRooms.Items.Clear();

            foreach (var room in _rooms.Where(r => !r.Disable).OrderBy(r => r.Sort))
            {
                var item = new RoomComboItem
                {
                    RoomId      = room.RoomId,
                    Name        = room.Name,
                    Remark      = room.Remark,
                    DisplayName = string.Format("{0} - {1}", room.RoomId, room.Name)
                };
                cmbRecRooms.Items.Add(item);
            }

            if (cmbRecRooms.Items.Count > 0)
                cmbRecRooms.SelectedIndex = 0;
        }

        private void CmbRooms_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedRoom = cmbRooms.SelectedItem as RoomComboItem;
            if (selectedRoom != null)
            {
                lblRemark.Text = selectedRoom.Remark;
            }
            RefreshAvailableSlots();
        }

        private async void DtpDate_ValueChanged(object sender, EventArgs e)
        {
            _selectedDate = dtpDate.Value;

            if (_fetchRecordsFunc != null)
            {
                try
                {
                    dgvAvailableSlots.Enabled = false;
                    this.Cursor = Cursors.WaitCursor;

                    var newRecords = await _fetchRecordsFunc(_selectedDate, _selectedDate);
                    if (newRecords != null)
                    {
                        _allRecords = newRecords;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[QuickBookingForm] 重新載入資料失敗: {ex.Message}");
                }
                finally
                {
                    dgvAvailableSlots.Enabled = true;
                    this.Cursor = Cursors.Default;
                }
            }

            RefreshAvailableSlots();
        }

        private void RefreshAvailableSlots()
        {
            if (cmbRooms.SelectedItem == null) return;

            var selectedRoom = cmbRooms.SelectedItem as RoomComboItem;
            if (selectedRoom == null) return;

            string roomId = selectedRoom.RoomId;
            DateTime date = dtpDate.Value.Date;

            dgvAvailableSlots.Rows.Clear();

            for (int hour = 8; hour <= 18; hour++)
            {
                AddTimeSlot(roomId, date, hour, 0);
                if (hour < 18)
                    AddTimeSlot(roomId, date, hour, 30);
            }

            btnBook.Enabled = false;
        }

        private void AddTimeSlot(string roomId, DateTime date, int hour, int minute)
        {
            DateTime slotStart = date.AddHours(hour).AddMinutes(minute);
            DateTime slotEnd   = slotStart.AddMinutes(30);

            bool isExpired   = date.Date == DateTime.Now.Date && slotStart < DateTime.Now;
            bool isAvailable = !isExpired && IsTimeSlotAvailable(roomId, slotStart, slotEnd);

            string timeRange = string.Format("{0:HH:mm} - {1:HH:mm}", slotStart, slotEnd);
            string duration  = "30分鐘";
            string bookerName = "";
            string subject    = "";
            string status;

            var booking = _allRecords.FirstOrDefault(r =>
                r.RoomId == roomId &&
                DateTime.Parse(r.StartDate).Date == date.Date &&
                slotStart < DateTime.Parse(r.EndDate) &&
                slotEnd   > DateTime.Parse(r.StartDate));

            if (booking != null)
            {
                bookerName = booking.UserName ?? "";
                subject    = booking.Subject  ?? "";
            }

            if (isExpired)
                status = "已逾時";
            else if (isAvailable)
                status = "可預約";
            else
                status = "已占用";

            int rowIndex = dgvAvailableSlots.Rows.Add(timeRange, status, bookerName, subject, duration);

            dgvAvailableSlots.Rows[rowIndex].Tag = new TimeSlotInfo
            {
                RoomId      = roomId,
                StartTime   = slotStart,
                EndTime     = slotEnd,
                IsAvailable = isAvailable && !isExpired
            };

            if (isExpired)
            {
                dgvAvailableSlots.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightGray;
                dgvAvailableSlots.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.Gray;
            }
            else if (isAvailable)
            {
                dgvAvailableSlots.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightGreen;
                dgvAvailableSlots.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.DarkGreen;
            }
            else
            {
                dgvAvailableSlots.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightCoral;
                dgvAvailableSlots.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.DarkRed;
            }
        }

        private bool IsTimeSlotAvailable(string roomId, DateTime start, DateTime end)
        {
            var roomBookings = _allRecords.Where(r =>
                r.RoomId == roomId &&
                DateTime.Parse(r.StartDate).Date == start.Date);

            foreach (var booking in roomBookings)
            {
                DateTime bookingStart = DateTime.Parse(booking.StartDate);
                DateTime bookingEnd   = DateTime.Parse(booking.EndDate);

                if (start < bookingEnd && end > bookingStart)
                    return false;
            }

            return true;
        }

        private void DgvAvailableSlots_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvAvailableSlots.SelectedRows.Count == 0)
            {
                btnBook.Enabled = false;
                return;
            }

            var selectedIndices = new List<int>();
            foreach (DataGridViewRow row in dgvAvailableSlots.SelectedRows)
                selectedIndices.Add(row.Index);
            selectedIndices.Sort();

            bool isContiguous = true;
            for (int i = 1; i < selectedIndices.Count; i++)
            {
                if (selectedIndices[i] - selectedIndices[i - 1] != 1)
                {
                    isContiguous = false;
                    break;
                }
            }

            bool allAvailable = true;
            foreach (int idx in selectedIndices)
            {
                var slotInfo = dgvAvailableSlots.Rows[idx].Tag as TimeSlotInfo;
                if (slotInfo == null || !slotInfo.IsAvailable)
                {
                    allAvailable = false;
                    break;
                }
            }

            if (isContiguous && allAvailable)
            {
                var firstSlot = dgvAvailableSlots.Rows[selectedIndices.First()].Tag as TimeSlotInfo;
                var lastSlot  = dgvAvailableSlots.Rows[selectedIndices.Last()].Tag  as TimeSlotInfo;

                _selectedRoomId   = firstSlot.RoomId;
                _selectedStartTime = firstSlot.StartTime;
                _selectedEndTime   = lastSlot.EndTime;

                var selectedRoom = cmbRooms.SelectedItem as RoomComboItem;
                if (selectedRoom != null)
                    SelectedRoomDisplayName = selectedRoom.DisplayName;

                btnBook.Enabled = true;
            }
            else
            {
                btnBook.Enabled = false;
            }
        }

        private void DgvAvailableSlots_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row      = dgvAvailableSlots.Rows[e.RowIndex];
            var slotInfo = row.Tag as TimeSlotInfo;

            if (slotInfo != null && slotInfo.IsAvailable)
            {
                _selectedRoomId    = slotInfo.RoomId;
                _selectedStartTime = slotInfo.StartTime;
                _selectedEndTime   = slotInfo.EndTime;

                var selectedRoom = cmbRooms.SelectedItem as RoomComboItem;
                if (selectedRoom != null)
                    SelectedRoomDisplayName = selectedRoom.DisplayName;

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private async void BtnBook_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_selectedRoomId))
            {
                var selectedRoom = cmbRooms.SelectedItem as RoomComboItem;
                if (selectedRoom != null)
                    SelectedRoomDisplayName = selectedRoom.DisplayName;

                using (var confirmForm = new Form())
                {
                    confirmForm.Text            = "確認預約";
                    confirmForm.Size            = new Size(450, 250);
                    confirmForm.StartPosition   = FormStartPosition.CenterParent;
                    confirmForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    confirmForm.MaximizeBox     = false;
                    confirmForm.MinimizeBox     = false;

                    var lblInfo = new Label();
                    lblInfo.Text = string.Format(
                        "會議室: {0}\n時間: {1:yyyy/MM/dd HH:mm} - {2:HH:mm}",
                        SelectedRoomDisplayName,
                        _selectedStartTime,
                        _selectedEndTime);
                    lblInfo.Location = new Point(20, 20);
                    lblInfo.Size     = new Size(400, 50);
                    lblInfo.Font     = new Font("Microsoft JhengHei", 10);
                    confirmForm.Controls.Add(lblInfo);

                    var lblSubject = new Label();
                    lblSubject.Text     = "會議主旨:";
                    lblSubject.Location = new Point(20, 80);
                    lblSubject.Size     = new Size(80, 25);
                    confirmForm.Controls.Add(lblSubject);

                    var txtSubject = new TextBox();
                    txtSubject.Location = new Point(110, 78);
                    txtSubject.Size     = new Size(300, 25);
                    confirmForm.Controls.Add(txtSubject);

                    var btnConfirm = new Button();
                    btnConfirm.Text         = "確認預約";
                    btnConfirm.DialogResult = DialogResult.Yes;
                    btnConfirm.Location     = new Point(230, 150);
                    btnConfirm.Size         = new Size(90, 30);
                    confirmForm.Controls.Add(btnConfirm);

                    var btnCancelConfirm = new Button();
                    btnCancelConfirm.Text         = "取消";
                    btnCancelConfirm.DialogResult = DialogResult.No;
                    btnCancelConfirm.Location     = new Point(330, 150);
                    btnCancelConfirm.Size         = new Size(80, 30);
                    confirmForm.Controls.Add(btnCancelConfirm);

                    confirmForm.AcceptButton = btnConfirm;
                    confirmForm.CancelButton = btnCancelConfirm;

                    var result = confirmForm.ShowDialog(this);

                    if (result == DialogResult.Yes)
                    {
                        MeetingSubject     = txtSubject.Text.Trim();
                        IsRecurrentBooking = false;

                        if (_bookRoomFunc != null)
                        {
                            this.Cursor     = Cursors.WaitCursor;
                            btnBook.Enabled = false;

                            try
                            {
                                bool bookingSuccess = await _bookRoomFunc(
                                    _selectedRoomId,
                                    SelectedRoomDisplayName,
                                    _selectedStartTime,
                                    _selectedEndTime,
                                    MeetingSubject);

                                if (bookingSuccess)
                                {
                                    MessageBox.Show(
                                        "會議室預約成功！\n即將開啟 Outlook 會議邀請。",
                                        "預約成功",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);
                                    this.DialogResult = DialogResult.OK;
                                    this.Close();
                                }
                                else
                                {
                                    MessageBox.Show(
                                        "會議室預約失敗，請重新選擇時段或稍後再試。",
                                        "預約失敗",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Warning);
                                    RefreshAvailableSlots();
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(
                                    $"預約時發生錯誤: {ex.Message}",
                                    "錯誤",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                            }
                            finally
                            {
                                this.Cursor     = Cursors.Default;
                                btnBook.Enabled = true;
                            }
                        }
                        else
                        {
                            this.DialogResult = DialogResult.OK;
                            this.Close();
                        }
                    }
                }
            }
        }

        #region 週期性預約功能

        private void BtnGeneratePreview_Click(object sender, EventArgs e)
        {
            GenerateRecurrencePreview();
        }

        private void GenerateRecurrencePreview()
        {
            var settings = GetRecurrenceSettings();
            if (settings == null) return;

            var dates = RecurrenceCalculator.CalculateDates(settings);

            dgvRecPreview.Rows.Clear();

            var selectedRoom = cmbRecRooms.SelectedItem as RoomComboItem;
            if (selectedRoom == null) return;

            int availableCount = 0;
            int occupiedCount  = 0;

            foreach (var date in dates)
            {
                DateTime slotStart  = date.Add(settings.StartTime);
                DateTime slotEnd    = date.Add(settings.EndTime);
                bool isAvailable    = IsTimeSlotAvailable(selectedRoom.RoomId, slotStart, slotEnd);
                string status       = isAvailable ? "可預約" : "已占用";

                if (isAvailable) availableCount++;
                else             occupiedCount++;

                int rowIndex = dgvRecPreview.Rows.Add(
                    date.ToString("yyyy/MM/dd (ddd)"),
                    string.Format("{0:HH:mm} - {1:HH:mm}", slotStart, slotEnd),
                    status,
                    selectedRoom.Name);

                if (isAvailable)
                {
                    dgvRecPreview.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightGreen;
                    dgvRecPreview.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.DarkGreen;
                }
                else
                {
                    dgvRecPreview.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightCoral;
                    dgvRecPreview.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.DarkRed;
                }
            }

            btnBookRecurrent.Enabled = availableCount > 0;

            if (occupiedCount > 0)
            {
                MessageBox.Show(
                    string.Format("預覽產生完成！\n可預約: {0} 個日期\n已占用: {1} 個日期\n\n已占用的日期將會被跳過。",
                        availableCount, occupiedCount),
                    "預覽結果",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(
                    string.Format("預覽產生完成！\n共 {0} 個日期可預約。", availableCount),
                    "預覽結果",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private RecurrenceSettings GetRecurrenceSettings()
        {
            var selectedRoom = cmbRecRooms.SelectedItem as RoomComboItem;
            if (selectedRoom == null)
            {
                MessageBox.Show("請選擇會議室。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            var daysOfWeek = new List<DayOfWeek>();
            if (cmbRecurrenceType.SelectedIndex == 1) // 每週
            {
                for (int i = 0; i < clbDaysOfWeek.Items.Count; i++)
                {
                    if (clbDaysOfWeek.GetItemChecked(i))
                        daysOfWeek.Add((DayOfWeek)((i + 1) % 7));
                }

                if (daysOfWeek.Count == 0)
                {
                    MessageBox.Show("請至少選擇一個星期幾。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }
            }

            var timeSlotParts = cmbTimeSlot.SelectedItem.ToString().Split('-');
            TimeSpan startTime = TimeSpan.Parse(timeSlotParts[0].Trim());
            TimeSpan endTime   = TimeSpan.Parse(timeSlotParts[1].Trim());

            var settings = new RecurrenceSettings
            {
                Type       = (RecurrenceType)cmbRecurrenceType.SelectedIndex,
                Interval   = (int)numInterval.Value,
                DaysOfWeek = daysOfWeek,
                StartDate  = dtpRecStartDate.Value,
                StartTime  = startTime,
                EndTime    = endTime
            };

            if (rdoEndByDate.Checked)
                settings.EndDate = dtpEndDate.Value;
            else
                settings.Occurrences = (int)numOccurrences.Value;

            return settings;
        }

        private async void BtnBookRecurrent_Click(object sender, EventArgs e)
        {
            var settings = GetRecurrenceSettings();
            if (settings == null) return;

            var selectedRoom = cmbRecRooms.SelectedItem as RoomComboItem;
            if (selectedRoom == null) return;

            var dates          = RecurrenceCalculator.CalculateDates(settings);
            var availableDates = new List<DateTime>();

            foreach (var date in dates)
            {
                DateTime slotStart = date.Add(settings.StartTime);
                DateTime slotEnd   = date.Add(settings.EndTime);
                if (IsTimeSlotAvailable(selectedRoom.RoomId, slotStart, slotEnd))
                    availableDates.Add(date);
            }

            if (availableDates.Count == 0)
            {
                MessageBox.Show("沒有可預約的日期。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var confirmForm = new Form())
            {
                confirmForm.Text            = "確認週期性預約";
                confirmForm.Size            = new Size(500, 350);
                confirmForm.StartPosition   = FormStartPosition.CenterParent;
                confirmForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                confirmForm.MaximizeBox     = false;
                confirmForm.MinimizeBox     = false;

                var lblInfo = new Label();
                lblInfo.Text = string.Format(
                    "會議室: {0}\n週期: {1}\n時段: {2}\n可預約日期數: {3}",
                    selectedRoom.Name,
                    GetRecurrenceDescription(settings),
                    cmbTimeSlot.SelectedItem,
                    availableDates.Count);
                lblInfo.Location = new Point(20, 20);
                lblInfo.Size     = new Size(450, 80);
                lblInfo.Font     = new Font("Microsoft JhengHei", 10);
                confirmForm.Controls.Add(lblInfo);

                var lblSubject = new Label();
                lblSubject.Text     = "會議主旨:";
                lblSubject.Location = new Point(20, 110);
                lblSubject.Size     = new Size(80, 25);
                confirmForm.Controls.Add(lblSubject);

                var txtSubject = new TextBox();
                txtSubject.Location = new Point(110, 108);
                txtSubject.Size     = new Size(350, 25);
                confirmForm.Controls.Add(txtSubject);

                var lblNote = new Label();
                lblNote.Text      = "注意：預約完成後將開啟 Outlook 週期性會議視窗。";
                lblNote.Location  = new Point(20, 150);
                lblNote.Size      = new Size(450, 25);
                lblNote.ForeColor = Color.Gray;
                lblNote.Font      = new Font("Microsoft JhengHei", 9, FontStyle.Italic);
                confirmForm.Controls.Add(lblNote);

                var btnConfirm = new Button();
                btnConfirm.Text         = "確認預約";
                btnConfirm.DialogResult = DialogResult.Yes;
                btnConfirm.Location     = new Point(280, 250);
                btnConfirm.Size         = new Size(90, 30);
                confirmForm.Controls.Add(btnConfirm);

                var btnCancelConfirm = new Button();
                btnCancelConfirm.Text         = "取消";
                btnCancelConfirm.DialogResult = DialogResult.No;
                btnCancelConfirm.Location     = new Point(380, 250);
                btnCancelConfirm.Size         = new Size(80, 30);
                confirmForm.Controls.Add(btnCancelConfirm);

                confirmForm.AcceptButton = btnConfirm;
                confirmForm.CancelButton = btnCancelConfirm;

                var result = confirmForm.ShowDialog(this);
                if (result == DialogResult.Yes)
                {
                    string subject = txtSubject.Text.Trim();
                    await ExecuteBatchBooking(selectedRoom, availableDates, settings, subject);
                }
            }
        }

        private string GetRecurrenceDescription(RecurrenceSettings settings)
        {
            string typeDesc = "";
            switch (settings.Type)
            {
                case RecurrenceType.Daily:
                    typeDesc = settings.Interval == 1 ? "每天" : $"每 {settings.Interval} 天";
                    break;
                case RecurrenceType.Weekly:
                    var days = string.Join(", ", settings.DaysOfWeek.Select(d =>
                        d == DayOfWeek.Monday    ? "一" :
                        d == DayOfWeek.Tuesday   ? "二" :
                        d == DayOfWeek.Wednesday ? "三" :
                        d == DayOfWeek.Thursday  ? "四" :
                        d == DayOfWeek.Friday    ? "五" :
                        d == DayOfWeek.Saturday  ? "六" : "日"));
                    typeDesc = settings.Interval == 1
                        ? $"每週 星期{days}"
                        : $"每 {settings.Interval} 週 星期{days}";
                    break;
                case RecurrenceType.Monthly:
                    typeDesc = settings.Interval == 1 ? "每月" : $"每 {settings.Interval} 月";
                    break;
            }
            return typeDesc;
        }

        private async Task ExecuteBatchBooking(RoomComboItem room, List<DateTime> dates,
            RecurrenceSettings settings, string subject)
        {
            if (_bookRoomFunc == null) return;

            this.Cursor              = Cursors.WaitCursor;
            btnBookRecurrent.Enabled = false;

            var result       = new BatchBookingResult();
            int successCount = 0;
            int failCount    = 0;

            try
            {
                foreach (var date in dates)
                {
                    DateTime slotStart = date.Add(settings.StartTime);
                    DateTime slotEnd   = date.Add(settings.EndTime);

                    bool success = await _bookRoomFunc(
                        room.RoomId, room.DisplayName,
                        slotStart, slotEnd, subject);

                    if (success)
                    {
                        successCount++;
                        result.SuccessfulBookings.Add(new BookingItem
                        {
                            Date   = date,
                            RoomId = room.RoomId
                        });
                    }
                    else
                    {
                        failCount++;
                        result.FailedBookings.Add(new FailedBookingItem
                        {
                            Date   = date,
                            Reason = "預約失敗"
                        });
                    }
                }

                result.Success    = failCount == 0;
                BatchBookingResult = result;

                if (failCount == 0)
                {
                    MessageBox.Show(
                        string.Format("批次預約成功！\n共預約 {0} 個日期。\n\n即將開啟 Outlook 週期性會議視窗。", successCount),
                        "預約成功",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    _selectedRoomId         = room.RoomId;
                    SelectedRoomDisplayName = room.DisplayName;
                    _selectedStartTime      = dates.First().Add(settings.StartTime);
                    _selectedEndTime        = dates.First().Add(settings.EndTime);
                    MeetingSubject          = subject;
                    IsRecurrentBooking      = true;
                    RecurrenceSettings      = settings;

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    var msg = string.Format(
                        "批次預約部分成功。\n成功: {0} 個日期\n失敗: {1} 個日期\n\n是否繼續開啟 Outlook 會議視窗？",
                        successCount, failCount);

                    var dialogResult = MessageBox.Show(msg, "預約結果", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.Yes)
                    {
                        _selectedRoomId         = room.RoomId;
                        SelectedRoomDisplayName = room.DisplayName;
                        _selectedStartTime      = dates.First().Add(settings.StartTime);
                        _selectedEndTime        = dates.First().Add(settings.EndTime);
                        MeetingSubject          = subject;
                        IsRecurrentBooking      = true;
                        RecurrenceSettings      = settings;

                        this.DialogResult = DialogResult.OK;
                        this.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"批次預約時發生錯誤: {ex.Message}",
                    "錯誤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor              = Cursors.Default;
                btnBookRecurrent.Enabled = true;
            }
        }

        #endregion

        private class TimeSlotInfo
        {
            public string   RoomId      { get; set; }
            public DateTime StartTime   { get; set; }
            public DateTime EndTime     { get; set; }
            public bool     IsAvailable { get; set; }
        }

        private class RoomComboItem
        {
            public string RoomId      { get; set; }
            public string Name        { get; set; }
            public string Remark      { get; set; }
            public string DisplayName { get; set; }

            public override string ToString() => DisplayName;
        }
    }
}