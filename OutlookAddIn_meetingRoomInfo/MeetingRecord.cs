public class MeetingRecord
{
    public string UserName { get; set; }
    public string RoomId { get; set; }
    public string StartDate { get; set; } // API 回傳的是字串格式
    public string EndDate { get; set; }
    public string Subject { get; set; }
    public string Remark { get; set; }
}