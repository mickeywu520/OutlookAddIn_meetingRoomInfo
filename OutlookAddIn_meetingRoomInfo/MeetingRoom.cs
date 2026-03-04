using System;

namespace OutlookAddIn_meetingRoomInfo
{
    /// <summary>
    /// 會議室資訊類別
    /// </summary>
    public class MeetingRoom
    {
        /// <summary>
        /// 會議室編號 (如 R001, R002)
        /// </summary>
        public string RoomId { get; set; }

        /// <summary>
        /// 會議室名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 會議室類型
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// 排序順序
        /// </summary>
        public int Sort { get; set; }

        /// <summary>
        /// 備註說明
        /// </summary>
        public string Remark { get; set; }

        /// <summary>
        /// 是否停用
        /// </summary>
        public bool Disable { get; set; }

        /// <summary>
        /// 取得顯示名稱
        /// </summary>
        public string DisplayName => string.Format("{0} - {1}", RoomId, Name);
    }
}
