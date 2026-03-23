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
        private const int DebounceMs = 500;

        private readonly ThisAddIn _addIn;

        public AppointmentMonitor(ThisAddIn addIn)
        {
            _addIn = addIn;
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

                if (_suppressedItems.Contains(entryId))
                {
                    System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] PropertyChange 被抑制: {propName}");
                    return;
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

            bool hasConflict = await CheckRoomAvailability(roomIdToCheck, newStart, newEnd, entryId);

            if (hasConflict)
            {
                var result = ShowConflictDialog(roomIdToCheck, newStart, newEnd);
                switch (result)
                {
                    case DialogResult.Yes:
                        RestoreOriginalTime(appointment, oldSnapshot);
                        break;
                    case DialogResult.No:
                        break;
                }
            }
            else
            {
                await RebookMeetingRoom(appointment, oldSnapshot, newRoomId, newStart, newEnd);
                UpdateSnapshot(appointment, newRoomId);
            }
        }

        private async Task<bool> CheckRoomAvailability(string roomId, DateTime start, DateTime end, string excludeEntryId)
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
                            return true;
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

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppointmentMonitor] CheckRoomAvailability 錯誤: {ex.Message}");
                return false;
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

        private DialogResult ShowConflictDialog(string roomId, DateTime start, DateTime end)
        {
            string message = $"會議室 {roomId} 在新時間段 ({start:HH:mm} - {end:HH:mm}) 已被預約！\n\n請選擇處理方式：";
            string caption = "會議室時間衝突";

            var dialog = new Form();
            dialog.Text = caption;
            dialog.Size = new Size(450, 220);
            dialog.StartPosition = FormStartPosition.CenterScreen;
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;

            var lblMessage = new Label();
            lblMessage.Text = message;
            lblMessage.Location = new Point(20, 20);
            lblMessage.Size = new Size(400, 60);
            dialog.Controls.Add(lblMessage);

            var btnRestore = new Button();
            btnRestore.Text = "還原為原時間";
            btnRestore.Location = new Point(20, 90);
            btnRestore.Size = new Size(130, 30);
            btnRestore.DialogResult = DialogResult.Yes;
            dialog.Controls.Add(btnRestore);

            var btnIgnore = new Button();
            btnIgnore.Text = "仍然儲存";
            btnIgnore.Location = new Point(160, 90);
            btnIgnore.Size = new Size(130, 30);
            btnIgnore.DialogResult = DialogResult.No;
            dialog.Controls.Add(btnIgnore);

            var btnCancel = new Button();
            btnCancel.Text = "取消變更";
            btnCancel.Location = new Point(300, 90);
            btnCancel.Size = new Size(110, 30);
            btnCancel.DialogResult = DialogResult.Cancel;
            dialog.Controls.Add(btnCancel);

            dialog.AcceptButton = btnRestore;
            dialog.CancelButton = btnCancel;

            return dialog.ShowDialog();
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
