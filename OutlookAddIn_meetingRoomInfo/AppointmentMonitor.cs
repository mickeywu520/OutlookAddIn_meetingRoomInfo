using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn_meetingRoomInfo
{
    public class AppointmentMonitor
    {
        private static readonly HttpClient client = new HttpClient();
        private readonly Dictionary<string, AppointmentSnapshot> _snapshots = new Dictionary<string, AppointmentSnapshot>();
        private readonly HashSet<string> _suppressedItems = new HashSet<string>();
        private readonly Dictionary<string, DateTime> _lastChangeTime = new Dictionary<string, DateTime>();
        private readonly Dictionary<string, DateTime> _lastUpdateFromBooking = new Dictionary<string, DateTime>();
        private const int DebounceMs = 500;
        private const int SelfUpdateWindowMs = 5000; // 5秒內視為自己更新（延長以確保回填觸發的非同步事件都能被攔截）

        private readonly ThisAddIn _addIn;

        public AppointmentMonitor(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        /// <summary>
        /// 標記即將從 QuickBookingForm 更新，之後的 PropertyChange 會被視為自己更新
        /// </summary>
        public void MarkUpdatingFromBooking(string entryId)
        {
            _suppressedItems.Add(entryId);
            _lastUpdateFromBooking[entryId] = DateTime.Now;
            System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] MarkUpdatingFromBooking: {entryId}");
        }

        /// <summary>
        /// 標記從 QuickBookingForm 更新完成，解除抑制
        /// </summary>
        public void ClearUpdatingFromBooking(string entryId)
        {
            _suppressedItems.Remove(entryId);
            System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] ClearUpdatingFromBooking: {entryId}");
        }

        public void RegisterInspector(Outlook.Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.AppointmentItem appointment)
            {
                RegisterAppointment(appointment);
            }
        }

        public void RegisterAppointment(Outlook.AppointmentItem appointment)
        {
            if (appointment == null) return;

            string entryId = appointment.EntryID ?? Guid.NewGuid().ToString();

            try
            {
                string roomId = GetRoomIdFromAppointment(appointment);
                var snapshot = new AppointmentSnapshot(appointment, roomId);
                _snapshots[entryId] = snapshot;

                appointment.PropertyChange += (propName) => Appointment_PropertyChange(appointment, propName);

                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已註冊監聽: EntryID={entryId.Substring(0, Math.Min(8, entryId.Length))}, RoomId={roomId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 註冊失敗: {ex.Message}");
            }
        }

        private void Appointment_PropertyChange(Outlook.AppointmentItem appointment, string propName)
        {
            if (propName == "Start" || propName == "End" || propName == "Location")
            {
                string entryId = appointment.EntryID;
                if (string.IsNullOrEmpty(entryId)) return;

                // 檢查是否為自己更新（從 QuickBookingForm）
                if (_suppressedItems.Contains(entryId))
                {
                    if (_lastUpdateFromBooking.TryGetValue(entryId, out DateTime lastUpdateTime))
                    {
                        if ((DateTime.Now - lastUpdateTime).TotalMilliseconds < SelfUpdateWindowMs)
                        {
                            // 2秒內的更新視為自己更新，跳過
                            System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 跳過自己更新: {propName}");
                            return;
                        }
                    }
                    
                    // 超過時間視為正常的 UI 變更，解除抑制
                    _suppressedItems.Remove(entryId);
                    _lastUpdateFromBooking.Remove(entryId);
                    System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 解除抑制（超時）: {entryId}");
                }

                DateTime now = DateTime.Now;
                if (_lastChangeTime.TryGetValue(entryId, out DateTime lastTime))
                {
                    if ((now - lastTime).TotalMilliseconds < DebounceMs)
                    {
                        System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] Debounce 忽略: {propName}");
                        return;
                    }
                }
                _lastChangeTime[entryId] = now;

                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] PropertyChange 觸發: {propName}");
                Task.Run(() => HandleTimeChange(appointment));
            }
        }

        private async void HandleTimeChange(Outlook.AppointmentItem appointment)
        {
            try
            {
                await HandleTimeChangeInternal(appointment);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] HandleTimeChange 錯誤: {ex.Message}");
            }
        }

        private void HandleTimeChangeSync(Outlook.AppointmentItem appointment)
        {
            try
            {
                HandleTimeChangeInternal(appointment).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] HandleTimeChangeSync 錯誤: {ex.Message}");
            }
        }

        private async Task HandleTimeChangeInternal(Outlook.AppointmentItem appointment)
        {
            string entryId = appointment.EntryID;
            if (string.IsNullOrEmpty(entryId)) return;

            if (!_snapshots.TryGetValue(entryId, out var oldSnapshot))
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 無快照記錄，跳過");
                return;
            }

            DateTime newStart = appointment.Start;
            DateTime newEnd = appointment.End;
            string newLocation = appointment.Location ?? "";

            if (!oldSnapshot.IsTimeChanged(newStart, newEnd) && !oldSnapshot.IsLocationChanged(newLocation))
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 時間和地點都沒有變更，跳過");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 偵測到變更:");
            System.Diagnostics.Debug.WriteLine($"  Old: {oldSnapshot.Start:yyyy/MM/dd HH:mm} - {oldSnapshot.End:HH:mm}, Location={oldSnapshot.Location}");
            System.Diagnostics.Debug.WriteLine($"  New: {newStart:yyyy/MM/dd HH:mm} - {newEnd:HH:mm}, Location={newLocation}");

            string newRoomId = ExtractRoomIdFromLocation(newLocation);
            string oldRoomId = oldSnapshot.RoomId;

            if (string.IsNullOrEmpty(newRoomId) && string.IsNullOrEmpty(oldRoomId))
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 非會議室預約，跳過");
                return;
            }

            string roomIdToCheck = !string.IsNullOrEmpty(newRoomId) ? newRoomId : oldRoomId;

            // 立即取消舊預約，避免自身時間衝突
            if (!string.IsNullOrEmpty(oldSnapshot.RoomId))
            {
                string userId = _addIn.GetCurrentUserId();
                string userName = _addIn.GetCurrentUserName();
                string userExt = _addIn.GetCurrentUserExt();

                string oldCaseId = await GetCaseIdFromRentRecord(oldSnapshot.RoomId, userId, oldSnapshot.Start, oldSnapshot.End);
                if (!string.IsNullOrEmpty(oldCaseId))
                {
                    await CancelBooking(oldCaseId, oldSnapshot.RoomId, userId, userName, oldSnapshot.Start, oldSnapshot.End, appointment.Subject, userExt);
                    System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已取消舊預約: CaseId={oldCaseId}");
                }
            }

            // 開啟選擇新時段對話框
            ShowRoomDetails(roomIdToCheck, newStart.Date, oldSnapshot);
        }

        /// <summary>
        /// 檢查會議室是否可用，返回衝突記錄列表
        /// </summary>
        private async Task<List<RentRecord>> CheckRoomAvailability(string roomId, DateTime start, DateTime end, string excludeEntryId)
        {
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/getRentRecord";

                System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] 檢查會議室: {roomId}");
                System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] 查詢時間範圍: {start:yyyy-MM-dd HH:mm} - {end:HH:mm}");
                System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] UTC時間: {start.ToUniversalTime():yyyy-MM-dd HH:mm:ss} - {end.ToUniversalTime():HH:mm:ss}");

                var payload = new
                {
                    CaseId = "",
                    RoomId = roomId,
                    UserId = "",
                    UserName = "",
                    StartDate = start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    EndDate = end.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    CreateTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    Subject = "",
                    Remark = "",
                    Cancel = false
                };

                string jsonPayload = JsonConvert.SerializeObject(payload);
                System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] Request JSON: {jsonPayload}");

                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] Response: {result}");

                    var records = JsonConvert.DeserializeObject<List<RentRecord>>(result);

                    if (records != null && records.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] API 回傳 {records.Count}筆記錄:");

                        // 過濾：只檢查相同 RoomId 且時間重疊的記錄
                        var roomRecords = records.Where(r => 
                            r.RoomId == roomId && 
                            !r.Cancel &&
                            IsTimeOverlap(start, end, r.StartDate, r.EndDate)).ToList();

                        System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] RoomId={roomId} 且時間重疊的記錄: {roomRecords.Count}筆");

                        foreach (var r in records)
                        {
                            bool overlaps = IsTimeOverlap(start, end, r.StartDate, r.EndDate);
                            System.Diagnostics.Debug.WriteLine($"  - CaseId={r.CaseId}, RoomId={r.RoomId}, Start={r.StartDate}, End={r.EndDate}, Cancel={r.Cancel}, 重疊={overlaps}");
                        }

                        if (roomRecords.Any())
                        {
                            System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 偵測到衝突: {roomRecords.Count} 筆預約");
                            foreach (var r in roomRecords)
                            {
                                System.Diagnostics.Debug.WriteLine($"  衝突: {r.UserName} - {r.Subject} ({r.StartDate} - {r.EndDate})");
                            }
                            return roomRecords;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] 會議室 {roomId} 在 {start:HH:mm}-{end:HH:mm} 可用");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] API 回傳空列表，會議室可用");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[CheckRoomAvailability] HTTP錯誤: {(int)response.StatusCode} {response.StatusCode}");
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] CheckRoomAvailability 錯誤: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 檢查兩個時間範圍是否重疊
        /// </summary>
        private bool IsTimeOverlap(DateTime start1, DateTime end1, string start2Str, string end2Str)
        {
            try
            {
                DateTime start2 = DateTime.Parse(start2Str);
                DateTime end2 = DateTime.Parse(end2Str);

                return start1 < end2 && end1 > start2;
            }
            catch
            {
                return false;
            }
        }

        private async Task RebookMeetingRoom(Outlook.AppointmentItem appointment, AppointmentSnapshot oldSnapshot, string newRoomId, DateTime newStart, DateTime newEnd)
        {
            try
            {
                string userId = _addIn.GetCurrentUserId();
                string userName = _addIn.GetCurrentUserName();
                string userExt = _addIn.GetCurrentUserExt();

                if (!string.IsNullOrEmpty(oldSnapshot.RoomId))
                {
                    string oldCaseId = await GetCaseIdFromRentRecord(oldSnapshot.RoomId, userId, oldSnapshot.Start, oldSnapshot.End);
                    if (!string.IsNullOrEmpty(oldCaseId))
                    {
                        await CancelBooking(oldCaseId, oldSnapshot.RoomId, userId, userName, oldSnapshot.Start, oldSnapshot.End, appointment.Subject, userExt);
                    }
                }

                if (!string.IsNullOrEmpty(newRoomId))
                {
                    string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/addRent";

                    var payload = new
                    {
                        CaseId = "",
                        RoomId = newRoomId,
                        UserId = userId,
                        UserName = userName,
                        StartDate = newStart.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                        EndDate = newEnd.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                        CreateTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                        Subject = appointment.Subject ?? "",
                        Remark = userExt,
                        Cancel = false
                    };

                    string jsonPayload = JsonConvert.SerializeObject(payload);
                    var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                    if (response.IsSuccessStatusCode)
                    {
                        string result = await response.Content.ReadAsStringAsync();
                        string cleanedResult = result.Trim().Trim('"');

                        if (cleanedResult == "1")
                        {
                            System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 重新預約成功");
                            MessageBox.Show(
                                $"會議室預約已更新！\n會議室: {newRoomId}\n新時間: {newStart:yyyy/MM/dd HH:mm} - {newEnd:HH:mm}",
                                "預約更新成功",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] RebookMeetingRoom 錯誤: {ex.Message}");
            }
        }

        private async Task<string> GetCaseIdFromRentRecord(string roomId, string userId, DateTime startDate, DateTime endDate)
        {
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/getRentRecord";

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
                    Cancel = false
                };

                string jsonPayload = JsonConvert.SerializeObject(payload);
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    var records = JsonConvert.DeserializeObject<List<RentRecord>>(result);

                    if (records != null)
                    {
                        foreach (var record in records)
                        {
                            if (record.RoomId == roomId && record.UserId == userId)
                            {
                                DateTime recordStart = DateTime.Parse(record.StartDate);
                                DateTime recordEnd = DateTime.Parse(record.EndDate);
                                TimeSpan startDiff = (recordStart - startDate).Duration();
                                TimeSpan endDiff = (recordEnd - endDate).Duration();

                                if (startDiff.TotalMinutes <= 5 && endDiff.TotalMinutes <= 5)
                                {
                                    return record.CaseId;
                                }
                            }
                        }

                        foreach (var record in records)
                        {
                            if (record.RoomId == roomId && record.UserId == userId)
                            {
                                return record.CaseId;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] GetCaseIdFromRentRecord 錯誤: {ex.Message}");
            }

            return null;
        }

        private async Task CancelBooking(string caseId, string roomId, string userId, string userName, DateTime start, DateTime end, string subject, string remark)
        {
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/editRent";

                var payload = new
                {
                    UserName = userName,
                    CaseId = caseId,
                    RoomId = roomId,
                    UserId = userId,
                    StartDate = start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    EndDate = end.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    CreateTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    Subject = subject ?? "",
                    Remark = remark,
                    Cancel = true
                };

                string jsonPayload = JsonConvert.SerializeObject(payload);
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                await client.PostAsync(apiUrl, content);
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已取消舊預約: CaseId={caseId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] CancelBooking 錯誤: {ex.Message}");
            }
        }

        private DialogResult ShowConflictDialog(string roomId, DateTime start, DateTime end, List<RentRecord> conflictRecords)
        {
            string message = "偵測到會議時間變更。\n舊預約已取消，請選擇新時段或還原。\n\n• 選擇新時段：查看當天預約並預約新時段\n• 還原為原時間：恢復原本的會議時間";
            string caption = "選擇處理方式";

            var dialog = new Form();
            dialog.Text = caption;
            dialog.Size = new Size(450, 250);
            dialog.StartPosition = FormStartPosition.CenterScreen;
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;

            var lblMessage = new Label();
            lblMessage.Text = message;
            lblMessage.Location = new Point(20, 20);
            lblMessage.Size = new Size(400, 100);
            dialog.Controls.Add(lblMessage);

            var btnNewTime = new Button();
            btnNewTime.Text = "選擇新時段";
            btnNewTime.Location = new Point(20, 130);
            btnNewTime.Size = new Size(130, 35);
            btnNewTime.DialogResult = DialogResult.No;
            dialog.Controls.Add(btnNewTime);

            var btnRestore = new Button();
            btnRestore.Text = "還原為原時間";
            btnRestore.Location = new Point(160, 130);
            btnRestore.Size = new Size(130, 35);
            btnRestore.DialogResult = DialogResult.Yes;
            dialog.Controls.Add(btnRestore);

            var btnCancel = new Button();
            btnCancel.Text = "取消變更";
            btnCancel.Location = new Point(300, 130);
            btnCancel.Size = new Size(110, 35);
            btnCancel.DialogResult = DialogResult.Cancel;
            dialog.Controls.Add(btnCancel);

            dialog.AcceptButton = btnNewTime;
            dialog.CancelButton = btnCancel;

            return dialog.ShowDialog();
        }

        private async void ShowRoomDetails(string roomId, DateTime date, AppointmentSnapshot snapshot)
        {
            try
            {
                var roomsTask = FetchMeetingRooms();
                var recordsTask = FetchRoomRecords(roomId, date);

                await Task.WhenAll(roomsTask, recordsTask);

                var rooms = await roomsTask;
                var records = await recordsTask;

                var meetingRecords = records.Select(r => new MeetingRecord
                {
                    UserName = r.UserName,
                    RoomId = r.RoomId,
                    StartDate = r.StartDate,
                    EndDate = r.EndDate,
                    Subject = r.Subject,
                    Remark = r.Remark
                }).ToList();

                using (var bookingForm = new QuickBookingForm(
                    meetingRecords, 
                    rooms, 
                    true,
                    snapshot.RoomId,
                    null,
                    snapshot.Start.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    snapshot.End.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    snapshot.Subject,
                    snapshot.EntryID,
                    null, null))
                {
                    bookingForm.SetSelectedDate(date);
                    bookingForm.SetSelectedRoom(roomId);

                    bookingForm.OnBookingUpdated = (newRoomId, newRoomDisplayName, newStart, newEnd) =>
                    {
                        UpdateAppointmentTime(snapshot, newRoomId, newRoomDisplayName, newStart, newEnd);
                    };

                    bookingForm.OnRestoreCompleted = () =>
                    {
                        RestoreOriginalTimeWithoutBooking(snapshot);
                    };

                    bookingForm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"無法取得預約資料: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RestoreOriginalTimeWithoutBooking(AppointmentSnapshot snapshot)
        {
            try
            {
                var session = _addIn.Application.Session;
                var calendarFolder = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                var items = calendarFolder.Items;

                Outlook.AppointmentItem targetAppointment = null;
                foreach (Outlook.AppointmentItem appt in items)
                {
                    if (appt.EntryID == snapshot.EntryID)
                    {
                        targetAppointment = appt;
                        break;
                    }
                }

                if (targetAppointment == null)
                {
                    MessageBox.Show("找不到對應的會議項目", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 先加入抑制，防止回填觸發重複事件
                _suppressedItems.Add(snapshot.EntryID);
                _lastUpdateFromBooking[snapshot.EntryID] = DateTime.Now;

                // 還原時間但不需要重新預約（因為已經取消了，現在只是還原時間）
                targetAppointment.Start = snapshot.Start;
                targetAppointment.End = snapshot.End;
                targetAppointment.Location = snapshot.Location;
                targetAppointment.Save();

                // 還原後更新快照，讓後續事件比對時判定「無變更」
                _snapshots[snapshot.EntryID] = new AppointmentSnapshot
                {
                    EntryID = snapshot.EntryID,
                    Start = snapshot.Start,
                    End = snapshot.End,
                    Location = snapshot.Location,
                    RoomId = snapshot.RoomId,
                    Subject = snapshot.Subject,
                    IsOrganizer = snapshot.IsOrganizer,
                    CapturedAt = DateTime.Now
                };

                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已還原時間並同步更新快照");

                // 延遲移除抑制
                var entryIdToRemove = snapshot.EntryID;
                Task.Delay(3000).ContinueWith(_ =>
                {
                    _suppressedItems.Remove(entryIdToRemove);
                    _lastUpdateFromBooking.Remove(entryIdToRemove);
                    System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 還原後已延遲移除抑制");
                });

                MessageBox.Show("已還原為原始時間。", "還原成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] RestoreOriginalTimeWithoutBooking 錯誤: {ex.Message}");
                MessageBox.Show($"還原失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateAppointmentTime(AppointmentSnapshot snapshot, string newRoomId, string newRoomDisplayName, DateTime newStart, DateTime newEnd)
        {
            try
            {
                var session = _addIn.Application.Session;
                var calendarFolder = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                var items = calendarFolder.Items;

                Outlook.AppointmentItem targetAppointment = null;
                foreach (Outlook.AppointmentItem appt in items)
                {
                    if (appt.EntryID == snapshot.EntryID)
                    {
                        targetAppointment = appt;
                        break;
                    }
                }

                if (targetAppointment == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 找不到Appointment: {snapshot.EntryID}");
                    MessageBox.Show("找不到對應的會議項目", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _suppressedItems.Add(snapshot.EntryID);
                _lastUpdateFromBooking[snapshot.EntryID] = DateTime.Now;

                targetAppointment.Start = newStart;
                targetAppointment.End = newEnd;
                targetAppointment.Location = newRoomDisplayName;
                targetAppointment.Save();

                // 回填後立即更新快照，讓後續 PropertyChange 比對時發現「無變更」而自然跳過
                _snapshots[snapshot.EntryID] = new AppointmentSnapshot
                {
                    EntryID = snapshot.EntryID,
                    Start = newStart,
                    End = newEnd,
                    Location = newRoomDisplayName,
                    RoomId = newRoomId,
                    Subject = snapshot.Subject,
                    IsOrganizer = snapshot.IsOrganizer,
                    CapturedAt = DateTime.Now
                };

                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已更新Appointment時間: {newStart} - {newEnd}");
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已同步更新快照，後續事件將判定無變更");

                // 延遲移除抑制，確保所有非同步事件（PropertyChange / CalendarItemChange）都被攔截
                var entryIdToRemove = snapshot.EntryID;
                Task.Delay(3000).ContinueWith(_ =>
                {
                    _suppressedItems.Remove(entryIdToRemove);
                    _lastUpdateFromBooking.Remove(entryIdToRemove);
                    System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已延遲移除抑制: {entryIdToRemove.Substring(0, Math.Min(8, entryIdToRemove.Length))}");
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] UpdateAppointmentTime錯誤: {ex.Message}");
                MessageBox.Show($"更新Outlook會議失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] FetchMeetingRooms 錯誤: {ex.Message}");
            }

            return GetDefaultRooms();
        }

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

        private async Task<List<RentRecord>> FetchRoomRecords(string roomId, DateTime date)
        {
            try
            {
                string apiUrl = "http://192.168.0.13:100/api/MeetingRoom/getRentRecord";

                var payload = new
                {
                    CaseId = "",
                    RoomId = roomId,
                    UserId = "",
                    UserName = "",
                    StartDate = date.Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    EndDate = date.Date.AddDays(1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    CreateTime = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
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
                    return JsonConvert.DeserializeObject<List<RentRecord>>(result) ?? new List<RentRecord>();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] FetchRoomRecords 錯誤: {ex.Message}");
            }

            return new List<RentRecord>();
        }

        private void RestoreOriginalTime(Outlook.AppointmentItem appointment, AppointmentSnapshot snapshot)
        {
            string entryId = appointment.EntryID;
            if (string.IsNullOrEmpty(entryId)) return;

            try
            {
                _suppressedItems.Add(entryId);

                appointment.Start = snapshot.Start;
                appointment.End = snapshot.End;
                appointment.Location = snapshot.Location;

                appointment.Save();

                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已還原時間");

                MessageBox.Show(
                    "已還原為原始時間。",
                    "還原成功",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] RestoreOriginalTime 錯誤: {ex.Message}");
            }
            finally
            {
                _suppressedItems.Remove(entryId);
            }
        }

        private void UpdateSnapshot(Outlook.AppointmentItem appointment, string roomId)
        {
            string entryId = appointment.EntryID;
            if (string.IsNullOrEmpty(entryId)) return;

            if (_snapshots.ContainsKey(entryId))
            {
                _snapshots[entryId] = new AppointmentSnapshot(appointment, roomId);
            }
        }

        private string GetRoomIdFromAppointment(Outlook.AppointmentItem appointment)
        {
            try
            {
                var userProps = appointment.UserProperties;
                if (userProps != null)
                {
                    var roomIdProp = userProps.Find("MeetingRoomId", false);
                    if (roomIdProp != null)
                    {
                        return roomIdProp.Value?.ToString();
                    }
                }
            }
            catch { }

            return ExtractRoomIdFromLocation(appointment.Location ?? "");
        }

        private string ExtractRoomIdFromLocation(string location)
        {
            if (string.IsNullOrEmpty(location))
                return null;

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

        public void UnregisterAppointment(string entryId)
        {
            if (_snapshots.ContainsKey(entryId))
            {
                _snapshots.Remove(entryId);
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] 已解除監聽: {entryId.Substring(0, Math.Min(8, entryId.Length))}");
            }
            _suppressedItems.Remove(entryId);
            _lastChangeTime.Remove(entryId);
        }

        public void HandleCalendarItemChange(Outlook.AppointmentItem appointment)
        {
            if (appointment == null) return;

            string entryId = appointment.EntryID;
            if (string.IsNullOrEmpty(entryId)) return;

            // 檢查是否為自己回填觸發的變更（封堵第二條觸發路徑）
            if (_suppressedItems.Contains(entryId))
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] CalendarItemChange 被抑制（回填觸發）: {entryId.Substring(0, Math.Min(8, entryId.Length))}");
                return;
            }

            if (!_snapshots.ContainsKey(entryId))
            {
                RegisterAppointment(appointment);
            }
            else
            {
                Task.Run(() => HandleTimeChange(appointment));
            }
        }
    }
}
