using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms; // 用於彈出視窗
using Newtonsoft.Json; // 需安裝 Newtonsoft.Json NuGet 套件
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn_meetingRoomInfo
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors _inspectors;
        private static readonly HttpClient client = new HttpClient();

        // Ribbon 執行個體
        private MeetingRoomRibbon _meetingRoomRibbon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _inspectors = this.Application.Inspectors;
            _inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            // 建立 Ribbon 執行個體
            _meetingRoomRibbon = new MeetingRoomRibbon();
        }

        private async void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            // 判斷是否為「新建立」的會議項目
            if (Inspector.CurrentItem is Outlook.AppointmentItem meetingItem)
            {
                if (string.IsNullOrEmpty(meetingItem.EntryID))
                {
                    // 開啟新會議時自動查詢當日租借狀況
                    await FetchAndShowMeetingRooms();
                }
            }
        }

        private async Task FetchAndShowMeetingRooms()
        {
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/getRentRecord";

                // 準備 POST Payload (根據你提供的 F12 內容)
                var payload = new
                {
                    StartDate = DateTime.Now.ToString("yyyy-MM-ddT00:00:00.000Z"),
                    EndDate = DateTime.Now.ToString("yyyy-MM-ddT23:59:59.000Z"),
                    // 其他欄位若 API 沒強制要求可帶空字串
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

                // 發送請求
                HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    var records = JsonConvert.DeserializeObject<List<MeetingRecord>>(result);

                    // 格式化顯示內容
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("=== 今日會議室預約狀況 ===");
                    foreach (var rec in records)
                    {
                        // 轉換時間格式，只顯示幾點幾分 (例如 16:30)
                        DateTime start = DateTime.Parse(rec.StartDate);
                        DateTime end = DateTime.Parse(rec.EndDate);
                        sb.AppendLine($"[{rec.RoomId}] {start:HH:mm}-{end:HH:mm} | {rec.UserName} ({rec.Subject})");
                    }

                    // 彈出對話框提示
                    MessageBox.Show(sb.ToString(), "會議室租借系統即時資訊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                // 內網 API 斷線或逾時處理
                System.Diagnostics.Debug.WriteLine("API 呼叫失敗: " + ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
