using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn_meetingRoomInfo
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors _inspectors;
        private static readonly HttpClient client = new HttpClient();

        // Ribbon 執行個體
        private MeetingRoomRibbon _meetingRoomRibbon;

        private void AccessContacts(string findLastName)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("1 -------- AccessContacts");
                Outlook.NameSpace session = this.Application.Session;

                Outlook.AddressEntry addrEntry =
                    session.CurrentUser.AddressEntry;

                if (addrEntry != null && addrEntry.Type == "EX")
                {
                    System.Diagnostics.Debug.WriteLine("2 -------- AccessContacts");
                    Outlook.ExchangeUser exchUser =
                        addrEntry.GetExchangeUser();

                    if (exchUser != null)
                    {
                        // 標準屬性
                        string name = exchUser.Name;
                        string email = exchUser.PrimarySmtpAddress;
                        string department = exchUser.Department;
                        string title = exchUser.JobTitle;

                        // 獲取「縮寫 (Initials)」欄位
                        // PR_INITIALS 的屬性架構 URL
                        string prInitials = "http://schemas.microsoft.com/mapi/proptag/0x3A0A001E";
                        string initials = "";

                        try
                        {
                            initials = (string)exchUser.PropertyAccessor.GetProperty(prInitials);
                        }
                        catch (Exception pEx)
                        {
                            initials = "無法獲取屬性: " + pEx.Message;
                        }

                        MessageBox.Show(
                            $"Name: {name}\n" +
                            $"Email: {email}\n" +
                            $"Title: {title}\n" +
                            $"Dept: {department}\n" +
                            $"Initials (縮寫): {initials}"); // 這裡應該會顯示 11754
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"4 -------- GetCurrentUserInfo error: {ex.Message}");
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("1 -------- ThisAddIn_Startup");
            _inspectors = this.Application.Inspectors;
            _inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            // 建立 Ribbon 執行個體
            _meetingRoomRibbon = new MeetingRoomRibbon();

            // 監聽會議項目的發送事件
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
            System.Diagnostics.Debug.WriteLine("2 -------- ThisAddIn_Startup");
            string windowsUserName = Environment.UserName;
            System.Diagnostics.Debug.WriteLine("3 -------- windowsUserName : " + windowsUserName);
            //string fullUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            //System.Diagnostics.Debug.WriteLine("4 -------- fullUserName : " + fullUserName);
        }

        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            // 判斷是否為「新建立」的會議項目
            if (Inspector.CurrentItem is Outlook.AppointmentItem meetingItem)
            {
                if (string.IsNullOrEmpty(meetingItem.EntryID))
                {
                    // 開啟新會議時的處理（目前無需自動查詢
                    //AccessContacts("mickey"); // test get current user details contacts
                }
            }
        }

        /// <summary>
        /// 當使用者發送會議邀請時觸發，自動預約會議室
        /// </summary>
        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            System.Diagnostics.Debug.WriteLine("[Application_ItemSend] ========== START ==========");
            System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] Item Type: {Item?.GetType()?.FullName ?? "null"}");
            
            try
            {
                // 首先嘗試透過 MessageClass 判斷項目類型
                string messageClass = "";
                try
                {
                    // 嘗試取得 MessageClass 屬性
                    dynamic dynamicItem = Item;
                    messageClass = dynamicItem.MessageClass;
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] MessageClass: {messageClass}");
                }
                catch (Exception mcEx)
                {
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] 無法取得 MessageClass: {mcEx.Message}");
                }
                
                // 檢查是否為會議邀請相關的 MessageClass
                // IPM.Appointment = 一般會議, IPM.Schedule.Meeting.Request = 會議邀請
                bool isMeetingRelated = messageClass.StartsWith("IPM.Appointment") || 
                                          messageClass.StartsWith("IPM.Schedule.Meeting");
                System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] 是否為會議相關: {isMeetingRelated}");
                
                if (!isMeetingRelated)
                {
                    System.Diagnostics.Debug.WriteLine("[Application_ItemSend] 非會議項目，跳過處理");
                    System.Diagnostics.Debug.WriteLine("[Application_ItemSend] ========== END ==========");
                    return;
                }
                
                // 嘗試將 Item 轉換為 AppointmentItem
                Outlook.AppointmentItem appointment = null;
                
                try
                {
                    // 使用反射或動態方式取得 AppointmentItem
                    dynamic dynamicItem = Item;
                    
                    // 如果是 MeetingItem，需要取得 GetAssociatedAppointment
                    if (messageClass.StartsWith("IPM.Schedule.Meeting"))
                    {
                        System.Diagnostics.Debug.WriteLine("[Application_ItemSend] 偵測到 MeetingItem，嘗試取得 Associated Appointment...");
                        try
                        {
                            appointment = dynamicItem.GetAssociatedAppointment(false);
                            System.Diagnostics.Debug.WriteLine("[Application_ItemSend] ✓ 成功取得 Associated Appointment");
                        }
                        catch (Exception gaEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] GetAssociatedAppointment 失敗: {gaEx.Message}");
                        }
                    }
                    else
                    {
                        // 直接嘗試轉換為 AppointmentItem
                        appointment = dynamicItem as Outlook.AppointmentItem;
                    }
                }
                catch (Exception castEx)
                {
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] 動態轉換失敗: {castEx.Message}");
                }
                
                // 如果上述方法失敗，嘗試標準轉換
                if (appointment == null)
                {
                    try
                    {
                        appointment = Item as Outlook.AppointmentItem;
                    }
                    catch (Exception castEx2)
                    {
                        System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] as 轉換失敗: {castEx2.Message}");
                    }
                }
                
                if (appointment != null)
                {
                    System.Diagnostics.Debug.WriteLine("[Application_ItemSend] ✓ 成功取得 AppointmentItem");
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] Subject: {appointment.Subject}");
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] Location: {appointment.Location}");
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] Start: {appointment.Start}");
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] End: {appointment.End}");
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] MeetingStatus: {appointment.MeetingStatus}");
                    
                    // 檢查是否為取消會議 (MeetingStatus = olMeetingCanceled = 5)
                    if (appointment.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled)
                    {
                        System.Diagnostics.Debug.WriteLine("[Application_ItemSend] 偵測到取消會議操作");
                        string roomId = ExtractRoomIdFromLocation(appointment.Location);
                        
                        if (!string.IsNullOrEmpty(roomId))
                        {
                            System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] 開始取消會議室預約，RoomId: {roomId}");
                            bool cancelSuccess = CancelMeetingRoomSync(appointment, roomId);
                            System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] CancelMeetingRoomSync 回傳: {cancelSuccess}");
                            
                            if (!cancelSuccess)
                            {
                                DialogResult result = MessageBox.Show(
                                    "會議室取消預約失敗，是否仍要取消會議？",
                                    "會議室取消警告",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Warning);
                                
                                if (result == DialogResult.No)
                                {
                                    Cancel = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        // 檢查是否已透過快速預約預約過會議室（主旨包含 [已預約] 標記）
                        bool isAlreadyBooked = !string.IsNullOrEmpty(appointment.Subject) && 
                                               appointment.Subject.StartsWith("[已預約] ");
                        
                        if (isAlreadyBooked)
                        {
                            System.Diagnostics.Debug.WriteLine("[Application_ItemSend] 偵測到 [已預約] 標記，跳過自動預約");
                            // 移除標記，還原原始主旨
                            appointment.Subject = appointment.Subject.Substring(8); // 移除 "[已預約] "
                        }
                        else
                        {
                            // 檢查是否有會議室位置（從 Location 欄位解析）
                            System.Diagnostics.Debug.WriteLine("[Application_ItemSend] 開始解析 Location...");
                            string roomId = ExtractRoomIdFromLocation(appointment.Location);
                            System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] 解析結果 RoomId: {roomId ?? "(null)"}");
                            
                            if (!string.IsNullOrEmpty(roomId))
                            {
                                System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] RoomId 有效，開始呼叫 BookMeetingRoomSync...");
                                // 使用同步方式處理預約（因為事件處理器需要立即決定是否取消發送）
                                bool bookingSuccess = BookMeetingRoomSync(appointment, roomId);
                                System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] BookMeetingRoomSync 回傳: {bookingSuccess}");
                                
                                if (!bookingSuccess)
                                {
                                    // 預約失敗，詢問使用者是否仍要發送會議邀請
                                    DialogResult result = MessageBox.Show(
                                        "會議室預約失敗，是否仍要發送會議邀請？",
                                        "會議室預約警告",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Warning);
                                    
                                    if (result == DialogResult.No)
                                    {
                                        Cancel = true; // 取消發送
                                    }
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine("[Application_ItemSend] RoomId 為空，跳過預約流程");
                            }
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] ✗ 無法取得 AppointmentItem");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] ✗ 發生例外錯誤: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[Application_ItemSend] StackTrace: {ex.StackTrace}");
            }
            
            System.Diagnostics.Debug.WriteLine("[Application_ItemSend] ========== END ==========");
        }

        /// <summary>
        /// 從會議地點解析會議室編號
        /// </summary>
        private string ExtractRoomIdFromLocation(string location)
        {
            if (string.IsNullOrEmpty(location))
                return null;

            // 支援的會議室關鍵字對應
            var roomMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "PARIS", "R001" },
                { "國際會議室", "R001" },
                { "TAIPEI", "R002" },
                { "大會議室", "R002" },
                { "SEOUL", "R003" },
                { "首爾", "R003" },
                { "SAN JOSE", "R005" },
                { "聖荷西", "R005" },
                { "LONDON", "R006" },
                { "業務會議室", "R006" },
                { "ZOOM", "R007" },
                { "達文西", "R008" },
                { "拉菲爾", "R009" },
                { "米開朗基羅", "R010" }
            };

            foreach (var mapping in roomMappings)
            {
                if (location.IndexOf(mapping.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return mapping.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// 同步方式取消會議室預約（用於取消會議時）
        /// </summary>
        private bool CancelMeetingRoomSync(Outlook.AppointmentItem appointment, string roomId)
        {
            System.Diagnostics.Debug.WriteLine("[CancelMeetingRoomSync] ========== 開始取消會議室預約 ==========");
            try
            {
                string userId = GetCurrentUserId();
                string userName = GetCurrentUserName();
                string userExt = GetCurrentUserExt(); // 取得分機號碼
                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] UserId: {userId}, UserName: {userName}, Ext: {userExt}");
                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] RoomId: {roomId}");
                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] Start: {appointment.Start}, End: {appointment.End}");

                // 步驟 1: 呼叫 getRentRecord 取得租借記錄以找到 CaseId
                string caseId = GetCaseIdFromRentRecord(roomId, userId, appointment.Start, appointment.End);
                
                if (string.IsNullOrEmpty(caseId))
                {
                    System.Diagnostics.Debug.WriteLine("[CancelMeetingRoomSync] ✗ 無法找到對應的租借記錄 (CaseId)");
                    MessageBox.Show(
                        "找不到會議室預約記錄，可能已經被取消或預約不存在。",
                        "取消預約失敗",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] 找到 CaseId: {caseId}");

                // 步驟 2: 呼叫 editRent 取消預約
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/editRent";
                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] API URL: {apiUrl}");

                var payload = new
                {
                    UserName = userName,
                    CaseId = caseId,
                    RoomId = roomId,
                    UserId = userId,
                    StartDate = appointment.Start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    EndDate = appointment.End.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    CreateTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    Subject = appointment.Subject ?? "",
                    Remark = userExt, // 使用分機號碼作為 Remark
                    Cancel = true
                };

                string jsonPayload = JsonConvert.SerializeObject(payload);
                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] Request JSON: {jsonPayload}");

                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                System.Diagnostics.Debug.WriteLine("[CancelMeetingRoomSync] 開始發送 HTTP POST 請求...");

                HttpResponseMessage response = client.PostAsync(apiUrl, content).GetAwaiter().GetResult();
                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] HTTP Status Code: {(int)response.StatusCode} {response.StatusCode}");

                if (response.IsSuccessStatusCode)
                {
                    string result = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] Response Body: {result}");

                    string cleanedResult = result.Trim().Trim('"');
                    System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] Cleaned Result: '{cleanedResult}'");

                    if (cleanedResult == "1")
                    {
                        System.Diagnostics.Debug.WriteLine("[CancelMeetingRoomSync] ✓ 取消預約成功！");
                        MessageBox.Show(
                            $"會議室取消預約成功！\n會議室: {roomId}\n時間: {appointment.Start:yyyy/MM/dd HH:mm} - {appointment.End:HH:mm}",
                            "取消預約成功",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        return true;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] ✗ API 回傳非成功代碼: {result}");
                    }
                }
                else
                {
                    string errorContent = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] ✗ HTTP 錯誤回應: {errorContent}");
                }

                System.Diagnostics.Debug.WriteLine("[CancelMeetingRoomSync] ========== 取消預約失敗 ==========");
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] ✗ 發生例外錯誤: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[CancelMeetingRoomSync] StackTrace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 從租借記錄中取得 CaseId
        /// </summary>
        private string GetCaseIdFromRentRecord(string roomId, string userId, DateTime startDate, DateTime endDate)
        {
            System.Diagnostics.Debug.WriteLine("[GetCaseIdFromRentRecord] ========== 開始查詢租借記錄 ==========");
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/getRentRecord";
                System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] API URL: {apiUrl}");

                // 準備查詢參數 - 使用較寬的時間範圍來確保能找到記錄
                var payload = new
                {
                    CaseId = "",
                    RoomId = roomId,
                    UserId = userId,
                    UserName = "",
                    StartDate = startDate.Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    EndDate = endDate.Date.AddDays(1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    CreateTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    Subject = "",
                    Remark = "",
                    Cancel = false,
                    MeetingRoom = new
                    {
                        RoomId = "",
                        Name = "",
                        Type = "",
                        Disable = false,
                        Remark = ""
                    }
                };

                string jsonPayload = JsonConvert.SerializeObject(payload);
                System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] Request JSON: {jsonPayload}");

                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                HttpResponseMessage response = client.PostAsync(apiUrl, content).GetAwaiter().GetResult();
                System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] HTTP Status Code: {(int)response.StatusCode} {response.StatusCode}");

                if (response.IsSuccessStatusCode)
                {
                    string result = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] Response Body: {result}");

                    // 解析 JSON 陣列
                    var records = JsonConvert.DeserializeObject<List<RentRecord>>(result);
                    
                    if (records != null && records.Count > 0)
                    {
                        // 尋找符合條件的記錄（相同房間、使用者、時間）
                        foreach (var record in records)
                        {
                            System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] 檢查記錄: CaseId={record.CaseId}, RoomId={record.RoomId}, UserId={record.UserId}, Start={record.StartDate}, End={record.EndDate}");
                            
                            // 比對房間和使用者
                            if (record.RoomId == roomId && record.UserId == userId)
                            {
                                // 比對時間（允許幾分鐘的誤差）
                                DateTime recordStart = DateTime.Parse(record.StartDate);
                                DateTime recordEnd = DateTime.Parse(record.EndDate);
                                
                                // 使用 TimeSpan 的 Duration() 取得絕對時間差
                                TimeSpan startDiff = (recordStart - startDate).Duration();
                                TimeSpan endDiff = (recordEnd - endDate).Duration();
                                
                                System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] 時間差異: startDiff={startDiff.TotalMinutes}min, endDiff={endDiff.TotalMinutes}min");
                                
                                // 如果時間差異在 5 分鐘內，視為同一筆預約
                                if (startDiff.TotalMinutes <= 5 && endDiff.TotalMinutes <= 5)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] ✓ 找到匹配的記錄，CaseId: {record.CaseId}");
                                    return record.CaseId;
                                }
                            }
                        }
                        
                        System.Diagnostics.Debug.WriteLine("[GetCaseIdFromRentRecord] ⚠ 未找到時間匹配的記錄，返回第一筆記錄的 CaseId");
                        // 如果沒有精確匹配，返回第一筆記錄（同房間、同使用者）
                        foreach (var record in records)
                        {
                            if (record.RoomId == roomId && record.UserId == userId)
                            {
                                return record.CaseId;
                            }
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[GetCaseIdFromRentRecord] ⚠ 查無租借記錄");
                    }
                }
                else
                {
                    string errorContent = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] ✗ HTTP 錯誤回應: {errorContent}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] ✗ 發生例外錯誤: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[GetCaseIdFromRentRecord] StackTrace: {ex.StackTrace}");
            }

            return null;
        }

        /// <summary>
        /// 同步方式預約會議室（用於事件處理器）
        /// </summary>
        private bool BookMeetingRoomSync(Outlook.AppointmentItem appointment, string roomId)
        {
            System.Diagnostics.Debug.WriteLine("[BookMeetingRoomSync] ========== 開始預約會議室 ==========");
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/addRent";
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] API URL: {apiUrl}");
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] RoomId: {roomId}");
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] Appointment Subject: {appointment.Subject}");
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] Appointment Start: {appointment.Start}");
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] Appointment End: {appointment.End}");
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] Appointment Location: {appointment.Location}");

                // 取得使用者資訊
                string userId = GetCurrentUserId();
                string userName = GetCurrentUserName();
                string userExt = GetCurrentUserExt(); // 取得分機號碼
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] UserId: {userId}, UserName: {userName}, Ext: {userExt}");

                // 準備 POST Payload
                var payload = new
                {
                    CaseId = "",
                    RoomId = roomId,
                    UserId = userId,
                    UserName = userName,
                    StartDate = appointment.Start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    EndDate = appointment.End.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    CreateTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    Subject = appointment.Subject ?? "",
                    Remark = userExt, // 使用分機號碼作為 Remark
                    Cancel = false,
                    MeetingRoom = new
                    {
                        RoomId = "",
                        Name = "",
                        Type = "",
                        Disable = false,
                        Remark = ""
                    }
                };

                string jsonPayload = JsonConvert.SerializeObject(payload);
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] Request JSON: {jsonPayload}");
                
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                System.Diagnostics.Debug.WriteLine("[BookMeetingRoomSync] 開始發送 HTTP POST 請求...");

                // 使用同步方式發送請求
                HttpResponseMessage response = client.PostAsync(apiUrl, content).GetAwaiter().GetResult();
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] HTTP Status Code: {(int)response.StatusCode} {response.StatusCode}");
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] IsSuccessStatusCode: {response.IsSuccessStatusCode}");

                if (response.IsSuccessStatusCode)
                {
                    string result = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] Response Body: {result}");
                    
                    // API 回傳 "1" 或 1 都表示成功（去除引號和空白後比較）
                    string cleanedResult = result.Trim().Trim('"');
                    System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] Cleaned Result: '{cleanedResult}'");
                    
                    if (cleanedResult == "1")
                    {
                        System.Diagnostics.Debug.WriteLine("[BookMeetingRoomSync] ✓ 預約成功！");
                        // 移除成功訊息，因為使用者已在 QuickBookingForm 中確認過預約
                        // 保留失敗時的錯誤處理
                        return true;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] ✗ API 回傳非成功代碼: {result}");
                    }
                }
                else
                {
                    // 記錄錯誤
                    string errorContent = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] ✗ HTTP 錯誤回應: {errorContent}");
                }
                
                System.Diagnostics.Debug.WriteLine("[BookMeetingRoomSync] ========== 預約失敗 ==========");
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] ✗ 發生例外錯誤: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[BookMeetingRoomSync] StackTrace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 取得目前使用者 ID（從 Outlook 使用者「縮寫」欄位讀取員工編號）
        /// </summary>
        private string GetCurrentUserId()
        {
            try
            {
                Outlook.NameSpace session = this.Application.Session;
                Outlook.AddressEntry addrEntry = session.CurrentUser.AddressEntry;

                if (addrEntry != null && addrEntry.Type == "EX")
                {
                    Outlook.ExchangeUser exchUser = addrEntry.GetExchangeUser();

                    if (exchUser != null)
                    {
                        // 獲取「縮寫 (Initials)」欄位 - 這是員工編號
                        // PR_INITIALS 的屬性架構 URL (0x3A0A001E 是正確的 property tag)
                        string prInitials = "http://schemas.microsoft.com/mapi/proptag/0x3A0A001E";
                        
                        try
                        {
                            string initials = (string)exchUser.PropertyAccessor.GetProperty(prInitials);
                            System.Diagnostics.Debug.WriteLine($"[GetCurrentUserId] 成功取得 Initials: {initials}");
                            
                            if (!string.IsNullOrEmpty(initials))
                                return initials.Trim();
                        }
                        catch (Exception pEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"[GetCurrentUserId] 取得 Initials 失敗: {pEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetCurrentUserId] 發生錯誤: {ex.Message}");
            }

            return ""; // 無法取得時回傳空字串
        }

        /// <summary>
        /// 取得目前使用者名稱
        /// </summary>
        private string GetCurrentUserName()
        {
            try
            {
                var currentUser = this.Application.Session.CurrentUser;
                if (currentUser != null)
                {
                    return currentUser.Name ?? "";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"取得使用者名稱失敗: {ex.Message}");
            }
            
            return "";
        }

        /// <summary>
        /// 取得目前使用者的分機號碼 (Ext) - 從 User API 查詢
        /// </summary>
        private string GetCurrentUserExt()
        {
            try
            {
                string userId = GetCurrentUserId();
                if (string.IsNullOrEmpty(userId))
                {
                    System.Diagnostics.Debug.WriteLine("[GetCurrentUserExt] UserId 為空，無法查詢分機號碼");
                    return $"磐儀#{userId}"; // fallback 使用 UserId
                }

                string apiUrl = "http://192.168.0.13:100/api/User/getAllUserListByEF";
                System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] 開始查詢使用者資訊，UserId: {userId}");

                HttpResponseMessage response = client.GetAsync(apiUrl).GetAwaiter().GetResult();
                System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] HTTP Status Code: {(int)response.StatusCode} {response.StatusCode}");

                if (response.IsSuccessStatusCode)
                {
                    string result = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] Response Body: {result.Substring(0, Math.Min(500, result.Length))}...");

                    // 解析 JSON
                    var userListResponse = JsonConvert.DeserializeObject<UserListResponse>(result);
                    
                    if (userListResponse?.Data != null)
                    {
                        // 尋找符合目前使用者的記錄
                        foreach (var user in userListResponse.Data)
                        {
                            if (user.UserId == userId)
                            {
                                System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] 找到使用者: {user.UserNameZH}, Ext: {user.Ext}");
                                
                                // 如果 Ext 不為空就使用它，否則 fallback 到磐儀#UserId
                                if (!string.IsNullOrEmpty(user.Ext))
                                {
                                    return user.Ext;
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] Ext 為空，使用 fallback: 磐儀#{userId}");
                                    return $"磐儀#{userId}";
                                }
                            }
                        }
                        
                        System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] 未找到使用者 {userId}，使用 fallback: 磐儀#{userId}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[GetCurrentUserExt] 無法解析使用者列表或資料為空");
                    }
                }
                else
                {
                    string errorContent = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] ✗ HTTP 錯誤回應: {errorContent}");
                }

                // fallback: 使用磐儀#UserId 格式
                return $"磐儀#{userId}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] ✗ 發生例外錯誤: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[GetCurrentUserExt] StackTrace: {ex.StackTrace}");
                return $"磐儀#{GetCurrentUserId()}";
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
