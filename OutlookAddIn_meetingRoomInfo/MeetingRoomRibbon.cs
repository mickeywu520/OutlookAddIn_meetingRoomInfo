using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn_meetingRoomInfo
{
    public partial class MeetingRoomRibbon : RibbonBase
    {
        private static readonly HttpClient client = new HttpClient();
        private List<MeetingRecord> _currentRecords = new List<MeetingRecord>();

        public MeetingRoomRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void MeetingRoomRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // 載入時可以初始化一些設定
        }

        /// <summary>
        /// 查詢今日會議室預約狀況
        /// </summary>
        private async void btnQueryToday_Click(object sender, RibbonControlEventArgs e)
        {
            await FetchAndShowMeetingRooms(DateTime.Now, DateTime.Now);
        }

        /// <summary>
        /// 查詢明日會議室預約狀況
        /// </summary>
        private async void btnQueryTomorrow_Click(object sender, RibbonControlEventArgs e)
        {
            var tomorrow = DateTime.Now.AddDays(1);
            await FetchAndShowMeetingRooms(tomorrow, tomorrow);
        }

        /// <summary>
        /// 選擇日期範圍查詢
        /// </summary>
        private async void btnQueryRange_Click(object sender, RibbonControlEventArgs e)
        {
            using (var form = new DateRangeForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    await FetchAndShowMeetingRooms(form.StartDate, form.EndDate);
                }
            }
        }

        /// <summary>
        /// 快速預約會議室 - 開啟新會議並帶入會議室資訊
        /// </summary>
        private void btnQuickBook_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var outlookApp = Globals.ThisAddIn.Application;
                var appointment = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);
                
                appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                appointment.Subject = "[會議室預約] ";
                appointment.Location = "";
                appointment.Start = DateTime.Now.AddHours(1).AddMinutes(-DateTime.Now.Minute).AddSeconds(-DateTime.Now.Second);
                appointment.Duration = 60;
                
                appointment.Display(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show("無法建立會議: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 顯示會議室使用說明
        /// </summary>
        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            StringBuilder help = new StringBuilder();
            help.AppendLine("=== 會議室管理系統使用說明 ===");
            help.AppendLine();
            help.AppendLine("【功能說明】");
            help.AppendLine("• 查詢今日：查看今天所有會議室的預約狀況");
            help.AppendLine("• 查詢明日：查看明天所有會議室的預約狀況");
            help.AppendLine("• 選擇日期：自訂日期範圍查詢預約記錄");
            help.AppendLine("• 快速預約：直接開啟新會議視窗進行預約");
            help.AppendLine();
            help.AppendLine("【注意事項】");
            help.AppendLine("• 請確保已連線至公司內網以取得即時資料");
            help.AppendLine("• 紅色時段表示該會議室已被預約");
            help.AppendLine("• 綠色時段表示該會議室為空閒狀態");
            
            MessageBox.Show(help.ToString(), "使用說明", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 呼叫 API 取得會議室預約資料
        /// </summary>
        private async Task FetchAndShowMeetingRooms(DateTime startDate, DateTime endDate)
        {
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/getRentRecord";

                var payload = new
                {
                    StartDate = startDate.ToString("yyyy-MM-ddT00:00:00.000Z"),
                    EndDate = endDate.ToString("yyyy-MM-ddT23:59:59.000Z"),
                    CaseId = "",
                    RoomId = "",
                    UserId = "",
                    UserName = "",
                    Subject = "",
                    Remark = "",
                    Cancel = false
                };

                string jsonPayload = JsonConvert.SerializeObject(payload);
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    _currentRecords = JsonConvert.DeserializeObject<List<MeetingRecord>>(result) ?? new List<MeetingRecord>();

                    ShowMeetingRoomResults(_currentRecords, startDate, endDate);
                }
                else
                {
                    MessageBox.Show($"API 回應錯誤: {response.StatusCode}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("無法取得會議室資料: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 顯示會議室查詢結果
        /// </summary>
        private void ShowMeetingRoomResults(List<MeetingRecord> records, DateTime startDate, DateTime endDate)
        {
            using (var resultForm = new MeetingRoomResultForm(records, startDate, endDate))
            {
                resultForm.ShowDialog();
            }
        }
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal MeetingRoomRibbon MeetingRoomRibbon
        {
            get { return this.GetRibbon<MeetingRoomRibbon>(); }
        }
    }
}
