using System;

namespace OutlookAddIn_meetingRoomInfo
{
    internal class RoomComboItem
    {
        public string Id { get; set; }
        public string Name { get; set; }

        public RoomComboItem() { }

        public RoomComboItem(string id, string name)
        {
            Id = id;
            Name = name;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
