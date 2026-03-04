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
        /// 快速預約會議室 - 顯示可用時段並開啟新會議
        /// </summary>
        private async void btnQuickBook_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 顯示載入提示
                var loadingForm = new Form();
                loadingForm.Text = "載入中...";
                loadingForm.Size = new Size(300, 100);
                loadingForm.StartPosition = FormStartPosition.CenterScreen;
                loadingForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                loadingForm.ControlBox = false;
                var lblLoading = new Label();
                lblLoading.Text = "正在取得會議室資訊...";
                lblLoading.Dock = DockStyle.Fill;
                lblLoading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                loadingForm.Controls.Add(lblLoading);
                loadingForm.Show();
                loadingForm.Refresh();

                try
                {
                    // 同時取得會議室清單和預約記錄
                    var roomsTask = FetchMeetingRooms();
                    var recordsTask = FetchMeetingRoomRecords(DateTime.Now, DateTime.Now.AddDays(7));

                    await Task.WhenAll(roomsTask, recordsTask);

                    var rooms = await roomsTask;
                    var records = await recordsTask;

                    loadingForm.Close();

                    using (var bookingForm = new QuickBookingForm(records, rooms, (start, end) => FetchMeetingRoomRecords(start, end)))
                    {
                        if (bookingForm.ShowDialog() == DialogResult.OK)
                        {
                            // 使用者選擇了時段，建立新會議
                            var outlookApp = Globals.ThisAddIn.Application;
                            var appointment = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);

                            string roomDisplayName = bookingForm.SelectedRoomDisplayName ?? bookingForm.SelectedRoomId;
                            
                            appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                            //appointment.Subject = string.Format("[會議室預約] {0}", roomDisplayName); // keeping subjetc null
                            appointment.Location = roomDisplayName;
                            appointment.Start = bookingForm.SelectedStartTime;
                            appointment.End = bookingForm.SelectedEndTime;

                            appointment.Display(false);
                        }
                    }
                }
                catch
                {
                    loadingForm.Close();
                    throw;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("無法建立會議: {0}", ex.Message), "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 取得會議室清單
        /// </summary>
        private async Task<List<MeetingRoom>> FetchMeetingRooms()
        {
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/getroomlist";

                HttpResponseMessage response = await client.GetAsync(apiUrl);

                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<List<MeetingRoom>>(result) ?? new List<MeetingRoom>();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("取得會議室清單失敗: {0}", ex.Message));
            }

            // Return default rooms if API fails
            return GetDefaultRooms();
        }

        /// <summary>
        /// 取得預設會議室清單（當 API 無法連線時使用）
        /// </summary>
        private List<MeetingRoom> GetDefaultRooms()
        {
            return new List<MeetingRoom>
            {
                new MeetingRoom { RoomId = "R001", Name = "PARIS(原國際會議室)", Sort = 1, Remark = "財務部旁", Disable = false },
                new MeetingRoom { RoomId = "R002", Name = "TAIPEI(原大會議室)", Sort = 2, Remark = "櫃檯後方大會議室", Disable = false },
                new MeetingRoom { RoomId = "R003", Name = "SEOUL(首爾會議室)", Sort = 3, Remark = "首爾會議室、軟體部前面", Disable = false },
                new MeetingRoom { RoomId = "R005", Name = "SAN JOSE(聖荷西會議室)", Sort = 5, Remark = "接待中心旁邊，5~6人", Disable = false },
                new MeetingRoom { RoomId = "R006", Name = "LONDON(原業務會議室)", Sort = 6, Remark = "業務區(可容納8-10人)", Disable = false },
                new MeetingRoom { RoomId = "R007", Name = "Zoom", Sort = 7, Remark = "Zoom 視訊會議室", Type = "虛擬", Disable = false },
                new MeetingRoom { RoomId = "R008", Name = "建康廠-達文西", Sort = 8, Remark = "4~6人", Type = "健康廠", Disable = false },
                new MeetingRoom { RoomId = "R009", Name = "建康廠-拉菲爾", Sort = 9, Remark = "4~6人", Type = "健康廠", Disable = false },
                new MeetingRoom { RoomId = "R010", Name = "建康廠-米開朗基羅", Sort = 10, Remark = "大會議室，12~15人", Type = "健康廠", Disable = false }
            };
        }

        /// <summary>
        /// 取得會議室預約記錄（供快速預約使用）
        /// </summary>
        private async Task<List<MeetingRecord>> FetchMeetingRoomRecords(DateTime startDate, DateTime endDate)
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
                    return JsonConvert.DeserializeObject<List<MeetingRecord>>(result) ?? new List<MeetingRecord>();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("API 呼叫失敗: {0}", ex.Message));
            }

            return new List<MeetingRecord>();
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
                // 同時取得會議室清單和預約記錄
                var roomsTask = FetchMeetingRooms();
                
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
                var rooms = await roomsTask;

                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    _currentRecords = JsonConvert.DeserializeObject<List<MeetingRecord>>(result) ?? new List<MeetingRecord>();

                    ShowMeetingRoomResults(_currentRecords, startDate, endDate, rooms);
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
        private void ShowMeetingRoomResults(List<MeetingRecord> records, DateTime startDate, DateTime endDate, List<MeetingRoom> rooms)
        {
            using (var resultForm = new MeetingRoomResultForm(records, startDate, endDate, rooms))
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
