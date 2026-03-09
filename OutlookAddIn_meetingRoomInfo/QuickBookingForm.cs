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

        // 用於立即預約會議室的委派
        private Func<string, string, DateTime, DateTime, string, Task<bool>> _bookRoomFunc;

        private ComboBox cmbRooms;
        private DateTimePicker dtpDate;
        private DataGridView dgvAvailableSlots;
        private Button btnBook;
        private Button btnCancel;
        private Label lblRoom;
        private Label lblDate;
        private Label lblTitle;
        private Label lblRemark;
        private Label lblLoading;

        public QuickBookingForm(List<MeetingRecord> existingRecords, List<MeetingRoom> rooms, 
            Func<DateTime, DateTime, Task<List<MeetingRecord>>> fetchRecordsFunc = null,
            Func<string, string, DateTime, DateTime, string, Task<bool>> bookRoomFunc = null)
        {
            _allRecords = existingRecords ?? new List<MeetingRecord>();
            _rooms = rooms ?? new List<MeetingRoom>();
            _fetchRecordsFunc = fetchRecordsFunc;
            _bookRoomFunc = bookRoomFunc;
            _selectedDate = DateTime.Now;
            InitializeComponent();
            LoadRooms();
            RefreshAvailableSlots();
        }

        private void InitializeComponent()
        {
            this.Text = "快速預約會議室";
            this.Size = new Size(850, 650);
            this.MinimumSize = new Size(800, 600);
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

            // Room label
            lblRoom = new Label();
            lblRoom.Text = "會議室:";
            lblRoom.Location = new Point(20, 55);
            lblRoom.Size = new Size(80, 25);
            this.Controls.Add(lblRoom);

            // Room combo box
            cmbRooms = new ComboBox();
            cmbRooms.Location = new Point(110, 53);
            cmbRooms.Size = new Size(280, 25);
            cmbRooms.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbRooms.SelectedIndexChanged += CmbRooms_SelectedIndexChanged;
            this.Controls.Add(cmbRooms);

            // Remark label (to show room info)
            lblRemark = new Label();
            lblRemark.Location = new Point(400, 55);
            lblRemark.Size = new Size(320, 25);
            lblRemark.Font = new Font("Microsoft JhengHei", 9, FontStyle.Italic);
            lblRemark.ForeColor = Color.Gray;
            this.Controls.Add(lblRemark);

            // Date label
            lblDate = new Label();
            lblDate.Text = "日期:";
            lblDate.Location = new Point(20, 90);
            lblDate.Size = new Size(50, 25);
            this.Controls.Add(lblDate);

            // Date picker
            dtpDate = new DateTimePicker();
            dtpDate.Location = new Point(110, 88);
            dtpDate.Size = new Size(150, 25);
            dtpDate.Format = DateTimePickerFormat.Short;
            dtpDate.MinDate = DateTime.Now.Date;
            dtpDate.ValueChanged += DtpDate_ValueChanged;
            this.Controls.Add(dtpDate);

            // Available slots grid
            dgvAvailableSlots = new DataGridView();
            dgvAvailableSlots.Location = new Point(20, 130);
            dgvAvailableSlots.Size = new Size(800, 430);
            dgvAvailableSlots.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvAvailableSlots.AllowUserToAddRows = false;
            dgvAvailableSlots.AllowUserToDeleteRows = false;
            dgvAvailableSlots.ReadOnly = true;
            dgvAvailableSlots.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvAvailableSlots.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvAvailableSlots.RowHeadersVisible = false;
            dgvAvailableSlots.MultiSelect = true;
            dgvAvailableSlots.SelectionChanged += DgvAvailableSlots_SelectionChanged;
            dgvAvailableSlots.CellDoubleClick += DgvAvailableSlots_CellDoubleClick;
            this.Controls.Add(dgvAvailableSlots);

            // Setup columns
            dgvAvailableSlots.Columns.Add("TimeSlot", "時段");
            dgvAvailableSlots.Columns.Add("Status", "狀態");
            dgvAvailableSlots.Columns.Add("Booker", "預約人");
            dgvAvailableSlots.Columns.Add("Subject", "會議主題");
            dgvAvailableSlots.Columns.Add("Duration", "時長");

            // Book button - 使用 Margin 方式定位，距離右邊和底部各 20px
            btnBook = new Button();
            btnBook.Text = "預約";
            btnBook.Size = new Size(100, 35);
            btnBook.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnBook.Enabled = false;
            btnBook.Click += BtnBook_Click;
            this.Controls.Add(btnBook);

            // Cancel button - 距離右邊 130px (20 + 100 + 10)，底部 20px
            btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Size = new Size(100, 35);
            btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.Controls.Add(btnCancel);

            // 在表單載入時設定按位置
            this.Load += (s, e) => {
                btnBook.Location = new Point(this.ClientSize.Width - 230, this.ClientSize.Height - 45);
                btnCancel.Location = new Point(this.ClientSize.Width - 120, this.ClientSize.Height - 45);
            };

            // // 在表單大小改變時更新按位置
            // this.Resize += (s, e) => {
            //     btnBook.Location = new Point(this.ClientSize.Width - 230, this.ClientSize.Height - 45);
            //     btnCancel.Location = new Point(this.ClientSize.Width - 120, this.ClientSize.Height - 45);
            // };

            this.AcceptButton = btnBook;
            this.CancelButton = btnCancel;
        }

        private void LoadRooms()
        {
            cmbRooms.Items.Clear();

            foreach (var room in _rooms.Where(r => !r.Disable).OrderBy(r => r.Sort))
            {
                var item = new RoomComboItem
                {
                    RoomId = room.RoomId,
                    Name = room.Name,
                    Remark = room.Remark,
                    DisplayName = string.Format("{0} - {1}", room.RoomId, room.Name)
                };
                cmbRooms.Items.Add(item);
            }

            if (cmbRooms.Items.Count > 0)
                cmbRooms.SelectedIndex = 0;
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
            
            // 如果有提供重新載入資料的委派，且選擇的日期不在初始資料範圍內，則重新載入
            if (_fetchRecordsFunc != null)
            {
                try
                {
                    // 顯示載入中提示
                    dgvAvailableSlots.Enabled = false;
                    this.Cursor = Cursors.WaitCursor;
                    
                    // 重新載入所選日期的資料
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

            // Generate time slots (8:00 - 18:30, every 30 minutes)
            for (int hour = 8; hour <= 18; hour++)
            {
                AddTimeSlot(roomId, date, hour, 0);
                if (hour < 18) // 18:00-18:30 為最後時段，不產生 18:30-19:00
                    AddTimeSlot(roomId, date, hour, 30);
            }

            btnBook.Enabled = false;
        }

        private void AddTimeSlot(string roomId, DateTime date, int hour, int minute)
        {
            DateTime slotStart = date.AddHours(hour).AddMinutes(minute);
            DateTime slotEnd = slotStart.AddMinutes(30);

            // 檢查是否為今天且時間已過期
            bool isExpired = date.Date == DateTime.Now.Date && slotStart < DateTime.Now;

            // Check if this slot is available (只檢查未過期的時段)
            bool isAvailable = !isExpired && IsTimeSlotAvailable(roomId, slotStart, slotEnd);

            string timeRange = string.Format("{0:HH:mm} - {1:HH:mm}", slotStart, slotEnd);
            string duration = "30分鐘";

            // 查找預約人名稱和會議主題（無論是否過期，都要查詢預約資訊）
            string bookerName = "";
            string subject = "";
            string status;

            // 先查詢是否有預約記錄
            var booking = _allRecords.FirstOrDefault(r =>
                r.RoomId == roomId &&
                DateTime.Parse(r.StartDate).Date == date.Date &&
                slotStart < DateTime.Parse(r.EndDate) && slotEnd > DateTime.Parse(r.StartDate));
            
            if (booking != null)
            {
                bookerName = booking.UserName ?? "";
                subject = booking.Subject ?? "";
            }

            if (isExpired)
            {
                status = "已逾時";
            }
            else if (isAvailable)
            {
                status = "可預約";
            }
            else
            {
                status = "已占用";
            }

            int rowIndex = dgvAvailableSlots.Rows.Add(timeRange, status, bookerName, subject, duration);

            // Store the actual datetime values in Tag for later use
            dgvAvailableSlots.Rows[rowIndex].Tag = new TimeSlotInfo
            {
                RoomId = roomId,
                StartTime = slotStart,
                EndTime = slotEnd,
                IsAvailable = isAvailable && !isExpired
            };

            // Color coding
            if (isExpired)
            {
                // 已逾時 - 灰色
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
            // Check against existing records
            var roomBookings = _allRecords.Where(r =>
                r.RoomId == roomId &&
                DateTime.Parse(r.StartDate).Date == start.Date);

            foreach (var booking in roomBookings)
            {
                DateTime bookingStart = DateTime.Parse(booking.StartDate);
                DateTime bookingEnd = DateTime.Parse(booking.EndDate);

                // Check for overlap
                if (start < bookingEnd && end > bookingStart)
                {
                    return false; // Overlapping, not available
                }
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

            // 取得所有選中行的索引並排序
            var selectedIndices = new List<int>();
            foreach (DataGridViewRow row in dgvAvailableSlots.SelectedRows)
            {
                selectedIndices.Add(row.Index);
            }
            selectedIndices.Sort();

            // 檢查是否連續
            bool isContiguous = true;
            for (int i = 1; i < selectedIndices.Count; i++)
            {
                if (selectedIndices[i] - selectedIndices[i - 1] != 1)
                {
                    isContiguous = false;
                    break;
                }
            }

            // 檢查所有選中的時段是否都為可預約
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
                // 取得首尾時段的時間
                var firstSlot = dgvAvailableSlots.Rows[selectedIndices.First()].Tag as TimeSlotInfo;
                var lastSlot = dgvAvailableSlots.Rows[selectedIndices.Last()].Tag as TimeSlotInfo;

                _selectedRoomId = firstSlot.RoomId;
                _selectedStartTime = firstSlot.StartTime;
                _selectedEndTime = lastSlot.EndTime;

                // Get the display name from current selection
                var selectedRoom = cmbRooms.SelectedItem as RoomComboItem;
                if (selectedRoom != null)
                {
                    SelectedRoomDisplayName = selectedRoom.DisplayName;
                }

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

            var row = dgvAvailableSlots.Rows[e.RowIndex];
            var slotInfo = row.Tag as TimeSlotInfo;

            if (slotInfo != null && slotInfo.IsAvailable)
            {
                _selectedRoomId = slotInfo.RoomId;
                _selectedStartTime = slotInfo.StartTime;
                _selectedEndTime = slotInfo.EndTime;
                
                // Get the display name from current selection
                var selectedRoom = cmbRooms.SelectedItem as RoomComboItem;
                if (selectedRoom != null)
                {
                    SelectedRoomDisplayName = selectedRoom.DisplayName;
                }
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private async void BtnBook_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_selectedRoomId))
            {
                // Get the display name from current selection
                var selectedRoom = cmbRooms.SelectedItem as RoomComboItem;
                if (selectedRoom != null)
                {
                    SelectedRoomDisplayName = selectedRoom.DisplayName;
                }

                // 建立自訂對話框，包含會議主旨輸入框
                using (var confirmForm = new Form())
                {
                    confirmForm.Text = "確認預約";
                    confirmForm.Size = new Size(450, 250);
                    confirmForm.StartPosition = FormStartPosition.CenterParent;
                    confirmForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    confirmForm.MaximizeBox = false;
                    confirmForm.MinimizeBox = false;

                    // 會議室資訊標籤
                    var lblInfo = new Label();
                    lblInfo.Text = string.Format(
                        "會議室: {0}\n時間: {1:yyyy/MM/dd HH:mm} - {2:HH:mm}",
                        SelectedRoomDisplayName,
                        _selectedStartTime,
                        _selectedEndTime);
                    lblInfo.Location = new Point(20, 20);
                    lblInfo.Size = new Size(400, 50);
                    lblInfo.Font = new Font("Microsoft JhengHei", 10);
                    confirmForm.Controls.Add(lblInfo);

                    // 會議主旨標籤
                    var lblSubject = new Label();
                    lblSubject.Text = "會議主旨:";
                    lblSubject.Location = new Point(20, 80);
                    lblSubject.Size = new Size(80, 25);
                    confirmForm.Controls.Add(lblSubject);

                    // 會議主旨輸入框
                    var txtSubject = new TextBox();
                    txtSubject.Location = new Point(110, 78);
                    txtSubject.Size = new Size(300, 25);
                    confirmForm.Controls.Add(txtSubject);

                    // 確認按
                    var btnConfirm = new Button();
                    btnConfirm.Text = "確認預約";
                    btnConfirm.DialogResult = DialogResult.Yes;
                    btnConfirm.Location = new Point(230, 150);
                    btnConfirm.Size = new Size(90, 30);
                    confirmForm.Controls.Add(btnConfirm);

                    // 取消按
                    var btnCancelConfirm = new Button();
                    btnCancelConfirm.Text = "取消";
                    btnCancelConfirm.DialogResult = DialogResult.No;
                    btnCancelConfirm.Location = new Point(330, 150);
                    btnCancelConfirm.Size = new Size(80, 30);
                    confirmForm.Controls.Add(btnCancelConfirm);

                    confirmForm.AcceptButton = btnConfirm;
                    confirmForm.CancelButton = btnCancelConfirm;

                    var result = confirmForm.ShowDialog(this);

                    if (result == DialogResult.Yes)
                    {
                        MeetingSubject = txtSubject.Text.Trim();

                        // 如果有提供預約委派，立即呼叫 API 預約
                        if (_bookRoomFunc != null)
                        {
                            this.Cursor = Cursors.WaitCursor;
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
                                    // 預約失敗，回到 ListView
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
                                this.Cursor = Cursors.Default;
                                btnBook.Enabled = true;
                            }
                        }
                        else
                        {
                            // 如果沒有提供預約委派，直接關閉表單（舊行為）
                            this.DialogResult = DialogResult.OK;
                            this.Close();
                        }
                    }
                    // 如果點擊「取消」或關閉對話框，則回到 ListView，不關閉表單
                }
            }
        }

        private class TimeSlotInfo
        {
            public string RoomId { get; set; }
            public DateTime StartTime { get; set; }
            public DateTime EndTime { get; set; }
            public bool IsAvailable { get; set; }
        }

        private class RoomComboItem
        {
            public string RoomId { get; set; }
            public string Name { get; set; }
            public string Remark { get; set; }
            public string DisplayName { get; set; }

            public override string ToString()
            {
                return DisplayName;
            }
        }
    }
}
