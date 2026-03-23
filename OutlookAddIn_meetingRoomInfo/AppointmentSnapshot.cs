using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

public class AppointmentSnapshot
{
    public string EntryID { get; set; }
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    public string Location { get; set; }
    public string RoomId { get; set; }
    public string Subject { get; set; }
    public bool IsOrganizer { get; set; }
    public DateTime CapturedAt { get; set; }

    public AppointmentSnapshot()
    {
        CapturedAt = DateTime.Now;
    }

    public AppointmentSnapshot(Outlook.AppointmentItem appointment, string roomId)
    {
        EntryID = appointment.EntryID;
        Start = appointment.Start;
        End = appointment.End;
        Location = appointment.Location ?? "";
        RoomId = roomId;
        Subject = appointment.Subject ?? "";
        try
        {
            IsOrganizer = appointment.MeetingStatus != Outlook.OlMeetingStatus.olMeetingReceived && appointment.MeetingStatus != Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled;
        }
        catch
        {
            IsOrganizer = true;
        }
        CapturedAt = DateTime.Now;
    }

    public bool IsTimeChanged(DateTime newStart, DateTime newEnd)
    {
        return Start != newStart || End != newEnd;
    }

    public bool IsLocationChanged(string newLocation)
    {
        return !string.Equals(Location, newLocation, StringComparison.OrdinalIgnoreCase);
    }
}

public class ConflictInfo
{
    public string RoomId { get; set; }
    public string RoomName { get; set; }
    public DateTime RequestedStart { get; set; }
    public DateTime RequestedEnd { get; set; }
    public string ExistingBooker { get; set; }
    public string ExistingSubject { get; set; }
}

public class RentRecord
{
    public string CaseId { get; set; }
    public string UserName { get; set; }
    public string RoomId { get; set; }
    public string UserId { get; set; }
    public string StartDate { get; set; }
    public string EndDate { get; set; }
    public string Subject { get; set; }
    public bool Cancel { get; set; }
}

public class MeetingRoom
{
    public string RoomId { get; set; }
    public string Name { get; set; }
    public string Type { get; set; }
    public int Sort { get; set; }
    public string Remark { get; set; }
    public bool Disable { get; set; }
}

public class MeetingRecord
{
    public string UserName { get; set; }
    public string RoomId { get; set; }
    public string StartDate { get; set; }
    public string EndDate { get; set; }
    public string Subject { get; set; }
    public string Remark { get; set; }
}

public class UserInfo
{
    public string UserId { get; set; }
    public string UserNameZH { get; set; }
    public string Ext { get; set; }
    public string Email { get; set; }
}

public class UserListResponse
{
    public string Code { get; set; }
    public string Message { get; set; }
    public List<UserInfo> Data { get; set; }
}
