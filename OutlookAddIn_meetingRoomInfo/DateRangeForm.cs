using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookAddIn_meetingRoomInfo
{
    public partial class DateRangeForm : Form
    {
        public DateTime StartDate { get; private set; }
        public DateTime EndDate { get; private set; }

        private DateTimePicker dtpStart;
        private DateTimePicker dtpEnd;
        private Button btnOK;
        private Button btnCancel;
        private Label lblStart;
        private Label lblEnd;

        public DateRangeForm()
        {
            InitializeComponent();
            StartDate = DateTime.Now;
            EndDate = DateTime.Now.AddDays(7);
        }

        private void InitializeComponent()
        {
            this.Text = "選擇查詢日期範圍";
            this.Size = new Size(350, 200);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 開始日期標籤
            lblStart = new Label();
            lblStart.Text = "開始日期:";
            lblStart.Location = new Point(20, 20);
            lblStart.Size = new Size(80, 25);
            this.Controls.Add(lblStart);

            // 開始日期選擇器
            dtpStart = new DateTimePicker();
            dtpStart.Location = new Point(110, 18);
            dtpStart.Size = new Size(200, 25);
            dtpStart.Format = DateTimePickerFormat.Short;
            dtpStart.Value = DateTime.Now;
            this.Controls.Add(dtpStart);

            // 結束日期標籤
            lblEnd = new Label();
            lblEnd.Text = "結束日期:";
            lblEnd.Location = new Point(20, 60);
            lblEnd.Size = new Size(80, 25);
            this.Controls.Add(lblEnd);

            // 結束日期選擇器
            dtpEnd = new DateTimePicker();
            dtpEnd.Location = new Point(110, 58);
            dtpEnd.Size = new Size(200, 25);
            dtpEnd.Format = DateTimePickerFormat.Short;
            dtpEnd.Value = DateTime.Now.AddDays(7);
            this.Controls.Add(dtpEnd);

            // 確定按鈕
            btnOK = new Button();
            btnOK.Text = "確定";
            btnOK.DialogResult = DialogResult.OK;
            btnOK.Location = new Point(130, 110);
            btnOK.Size = new Size(80, 30);
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // 取消按鈕
            btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(230, 110);
            btnCancel.Size = new Size(80, 30);
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (dtpEnd.Value < dtpStart.Value)
            {
                MessageBox.Show("結束日期不能早於開始日期！", "日期錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }

            StartDate = dtpStart.Value;
            EndDate = dtpEnd.Value;
        }
    }
}
