using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookAddIn_meetingRoomInfo
{
    public partial class MeetingRoomResultForm : Form
    {
        private List<MeetingRecord> _records;
        private List<MeetingRoom> _rooms;
        private DateTime _startDate;
        private DateTime _endDate;
        private DataGridView dgvResults;
        private Button btnClose;
        private Button btnExport;
        private Label lblTitle;
        private Label lblSummary;

        public MeetingRoomResultForm(List<MeetingRecord> records, DateTime startDate, DateTime endDate, List<MeetingRoom> rooms = null)
        {
            _records = records ?? new List<MeetingRecord>();
            _rooms = rooms ?? new List<MeetingRoom>();
            _startDate = startDate;
            _endDate = endDate;
            InitializeComponent();
            LoadData();
        }

        private void InitializeComponent()
        {
            this.Text = string.Format("會議室預約查詢結果 ({0:yyyy/MM/dd} - {1:yyyy/MM/dd})", _startDate, _endDate);
            this.Size = new Size(900, 600);
            this.MinimumSize = new Size(700, 400);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Title label
            lblTitle = new Label();
            lblTitle.Text = "會議室預約狀況查詢結果";
            lblTitle.Font = new Font("Microsoft JhengHei", 14, FontStyle.Bold);
            lblTitle.Location = new Point(20, 15);
            lblTitle.Size = new Size(400, 30);
            this.Controls.Add(lblTitle);

            // Summary label
            lblSummary = new Label();
            lblSummary.Location = new Point(20, 50);
            lblSummary.Size = new Size(600, 25);
            lblSummary.Font = new Font("Microsoft JhengHei", 10);
            this.Controls.Add(lblSummary);

            // DataGridView
            dgvResults = new DataGridView();
            dgvResults.Location = new Point(20, 85);
            dgvResults.Size = new Size(840, 420);
            dgvResults.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvResults.AllowUserToAddRows = false;
            dgvResults.AllowUserToDeleteRows = false;
            dgvResults.ReadOnly = true;
            dgvResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvResults.RowHeadersVisible = false;
            dgvResults.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
            dgvResults.CellDoubleClick += DgvResults_CellDoubleClick;
            this.Controls.Add(dgvResults);

            // Close button
            btnClose = new Button();
            btnClose.Text = "關閉";
            btnClose.DialogResult = DialogResult.OK;
            btnClose.Location = new Point(780, 520);
            btnClose.Size = new Size(80, 30);
            btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.Controls.Add(btnClose);

            // Export button
            btnExport = new Button();
            btnExport.Text = "匯出 CSV";
            btnExport.Location = new Point(680, 520);
            btnExport.Size = new Size(90, 30);
            btnExport.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnExport.Click += BtnExport_Click;
            this.Controls.Add(btnExport);

            this.AcceptButton = btnClose;
        }

        /// <summary>
        /// 取得會議室顯示名稱（包含 RoomId 和 Name）
        /// </summary>
        private string GetRoomDisplayName(string roomId)
        {
            var room = _rooms.FirstOrDefault(r => r.RoomId == roomId);
            if (room != null)
            {
                return $"{room.RoomId} - {room.Name}";
            }
            return roomId; // 如果找不到對應的會議室，只回傳 RoomId
        }

        private void LoadData()
        {
            // Setup columns
            dgvResults.Columns.Add("RoomId", "會議室");
            dgvResults.Columns.Add("StartDate", "開始時間");
            dgvResults.Columns.Add("EndDate", "結束時間");
            dgvResults.Columns.Add("UserName", "預約人");
            dgvResults.Columns.Add("Subject", "會議主題");
            dgvResults.Columns.Add("Remark", "備註");

            // Set column widths
            dgvResults.Columns["RoomId"].FillWeight = 20;
            dgvResults.Columns["StartDate"].FillWeight = 16;
            dgvResults.Columns["EndDate"].FillWeight = 16;
            dgvResults.Columns["UserName"].FillWeight = 14;
            dgvResults.Columns["Subject"].FillWeight = 22;
            dgvResults.Columns["Remark"].FillWeight = 12;

            // Sort records
            var sortedRecords = _records.OrderBy(r => r.RoomId)
                                       .ThenBy(r => DateTime.Parse(r.StartDate))
                                       .ToList();

            // Fill data
            foreach (var record in sortedRecords)
            {
                DateTime start = DateTime.Parse(record.StartDate);
                DateTime end = DateTime.Parse(record.EndDate);

                // 取得會議室顯示名稱（ID + 名稱）
                string roomDisplayName = GetRoomDisplayName(record.RoomId);

                int rowIndex = dgvResults.Rows.Add(
                    roomDisplayName,
                    start.ToString("yyyy/MM/dd HH:mm"),
                    end.ToString("yyyy/MM/dd HH:mm"),
                    record.UserName,
                    record.Subject,
                    record.Remark
                );

                // Set color based on date
                if (start.Date == DateTime.Now.Date)
                {
                    dgvResults.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
                }
            }

            // Update summary
            int totalBookings = _records.Count;
            var distinctRooms = _records.Select(r => r.RoomId).Distinct().Count();
            lblSummary.Text = string.Format("總計: {0} 筆預約記錄 | 涉及 {1} 間會議室", totalBookings, distinctRooms);
        }

        private void DgvResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = dgvResults.Rows[e.RowIndex];
            StringBuilder details = new StringBuilder();
            details.AppendLine(string.Format("會議室: {0}", row.Cells["RoomId"].Value));
            details.AppendLine(string.Format("開始時間: {0}", row.Cells["StartDate"].Value));
            details.AppendLine(string.Format("結束時間: {0}", row.Cells["EndDate"].Value));
            details.AppendLine(string.Format("預約人: {0}", row.Cells["UserName"].Value));
            details.AppendLine(string.Format("會議主題: {0}", row.Cells["Subject"].Value));
            details.AppendLine(string.Format("備註: {0}", row.Cells["Remark"].Value));

            MessageBox.Show(details.ToString(), "預約詳細資訊", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                saveDialog.FileName = string.Format("MeetingRoom_Booking_{0:yyyyMMdd}_{1:yyyyMMdd}.csv", _startDate, _endDate);

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    StringBuilder csv = new StringBuilder();
                    csv.AppendLine("Room,Start Time,End Time,Booked By,Subject,Remark");

                    foreach (DataGridViewRow row in dgvResults.Rows)
                    {
                        if (row.IsNewRow) continue;
                        csv.AppendLine(string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\"",
                            row.Cells["RoomId"].Value,
                            row.Cells["StartDate"].Value,
                            row.Cells["EndDate"].Value,
                            row.Cells["UserName"].Value,
                            row.Cells["Subject"].Value,
                            row.Cells["Remark"].Value));
                    }

                    File.WriteAllText(saveDialog.FileName, csv.ToString(), Encoding.UTF8);
                    MessageBox.Show("Export successful!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Export failed: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
