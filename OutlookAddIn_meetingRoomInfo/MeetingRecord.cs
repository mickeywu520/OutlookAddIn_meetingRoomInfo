using System.Collections.Generic;

namespace OutlookAddIn_meetingRoomInfo
{
    public class MeetingRecord
    {
        public string UserName { get; set; }
        public string RoomId { get; set; }
        public string StartDate { get; set; } // API 回傳的是字串格式
        public string EndDate { get; set; }
        public string Subject { get; set; }
        public string Remark { get; set; }
    }

    /// <summary>
    /// 租借記錄類別 - 用於解析 getRentRecord API 回傳的 JSON
    /// </summary>
    public class RentRecord
    {
        public string CaseId { get; set; }
        public string UserName { get; set; }
        public string RoomId { get; set; }
        public string UserId { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string CreateTime { get; set; }
        public string Subject { get; set; }
        public string Remark { get; set; }
        public bool Cancel { get; set; }
    }

    /// <summary>
    /// 使用者資訊 API 回應類別
    /// </summary>
    public class UserListResponse
    {
        public string Code { get; set; }
        public string Message { get; set; }
        public List<UserInfo> Data { get; set; }
    }

    /// <summary>
    /// 使用者資訊類別
    /// </summary>
    public class UserInfo
    {
        public string UserId { get; set; }
        public string CompanyId { get; set; }
        public string UserNameZH { get; set; }
        public string UserNameEN { get; set; }
        public string ENName { get; set; }
        public string DepartmentId { get; set; }
        public string DepartmentName { get; set; }
        public int DepartmentSort { get; set; }
        public string JobTitleId { get; set; }
        public string Ext { get; set; }
        public string Email { get; set; }
    }
}
