using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookAddIn_meetingRoomInfo
{
    /// <summary>
    /// 週期性會議設定類別
    /// </summary>
    public class RecurrenceSettings
    {
        /// <summary>
        /// 週期類型
        /// </summary>
        public RecurrenceType Type { get; set; }

        /// <summary>
        /// 間隔（每幾天/幾週/幾月）
        /// </summary>
        public int Interval { get; set; }

        /// <summary>
        /// 星期幾（用於每週週期）
        /// </summary>
        public List<DayOfWeek> DaysOfWeek { get; set; }

        /// <summary>
        /// 開始日期
        /// </summary>
        public DateTime StartDate { get; set; }

        /// <summary>
        /// 結束日期（可為 null）
        /// </summary>
        public DateTime? EndDate { get; set; }

        /// <summary>
        /// 重複次數（可為 null）
        /// </summary>
        public int? Occurrences { get; set; }

        /// <summary>
        /// 固定時段開始時間
        /// </summary>
        public TimeSpan StartTime { get; set; }

        /// <summary>
        /// 固定時段結束時間
        /// </summary>
        public TimeSpan EndTime { get; set; }

        public RecurrenceSettings()
        {
            DaysOfWeek = new List<DayOfWeek>();
            Interval = 1;
            Type = RecurrenceType.Weekly;
        }
    }

    /// <summary>
    /// 週期類型列舉
    /// </summary>
    public enum RecurrenceType
    {
        Daily,
        Weekly,
        Monthly
    }

    /// <summary>
    /// 週期日期計算器
    /// </summary>
    public static class RecurrenceCalculator
    {
        /// <summary>
        /// 計算所有週期日期
        /// </summary>
        public static List<DateTime> CalculateDates(RecurrenceSettings settings)
        {
            var dates = new List<DateTime>();
            var currentDate = settings.StartDate.Date;
            var endDate = settings.EndDate;
            var maxOccurrences = settings.Occurrences;

            // 如果沒有設定結束條件，預設 10 次避免無限迴圈
            if (!endDate.HasValue && !maxOccurrences.HasValue)
            {
                maxOccurrences = 10;
            }

            switch (settings.Type)
            {
                case RecurrenceType.Daily:
                    CalculateDailyRecurrence(dates, currentDate, endDate, maxOccurrences, settings.Interval);
                    break;

                case RecurrenceType.Weekly:
                    CalculateWeeklyRecurrence(dates, currentDate, endDate, maxOccurrences, settings.Interval, settings.DaysOfWeek);
                    break;

                case RecurrenceType.Monthly:
                    CalculateMonthlyRecurrence(dates, currentDate, endDate, maxOccurrences, settings.Interval);
                    break;
            }

            return dates.OrderBy(d => d).ToList();
        }

        private static void CalculateDailyRecurrence(List<DateTime> dates, DateTime startDate, DateTime? endDate, int? maxOccurrences, int interval)
        {
            var currentDate = startDate;
            int count = 0;

            while (true)
            {
                if (endDate.HasValue && currentDate > endDate.Value)
                    break;

                if (maxOccurrences.HasValue && count >= maxOccurrences.Value)
                    break;

                dates.Add(currentDate);
                currentDate = currentDate.AddDays(interval);
                count++;

                // 安全機制：最多計算 100 個日期
                if (count >= 100)
                    break;
            }
        }

        private static void CalculateWeeklyRecurrence(List<DateTime> dates, DateTime startDate, DateTime? endDate, int? maxOccurrences, int interval, List<DayOfWeek> daysOfWeek)
        {
            if (daysOfWeek == null || daysOfWeek.Count == 0)
            {
                // 如果沒有選擇星期幾，預設使用開始日期的星期
                daysOfWeek = new List<DayOfWeek> { startDate.DayOfWeek };
            }

            // 排序星期幾
            daysOfWeek = daysOfWeek.OrderBy(d => d).ToList();

            var currentWeekStart = GetWeekStart(startDate);
            int count = 0;

            while (true)
            {
                foreach (var dayOfWeek in daysOfWeek)
                {
                    if (maxOccurrences.HasValue && count >= maxOccurrences.Value)
                        break;

                    var targetDate = GetDateOfWeek(currentWeekStart, dayOfWeek);

                    // 確保日期不早於開始日期
                    if (targetDate < startDate.Date)
                        continue;

                    if (endDate.HasValue && targetDate > endDate.Value)
                        break;

                    dates.Add(targetDate);
                    count++;
                }

                if (maxOccurrences.HasValue && count >= maxOccurrences.Value)
                    break;

                if (endDate.HasValue && currentWeekStart.AddDays(7 * interval) > endDate.Value.AddDays(7))
                    break;

                currentWeekStart = currentWeekStart.AddDays(7 * interval);

                // 安全機制：最多計算 100 個日期
                if (count >= 100)
                    break;
            }
        }

        private static void CalculateMonthlyRecurrence(List<DateTime> dates, DateTime startDate, DateTime? endDate, int? maxOccurrences, int interval)
        {
            var currentDate = startDate;
            int count = 0;

            while (true)
            {
                if (endDate.HasValue && currentDate > endDate.Value)
                    break;

                if (maxOccurrences.HasValue && count >= maxOccurrences.Value)
                    break;

                dates.Add(currentDate);
                currentDate = currentDate.AddMonths(interval);
                count++;

                // 安全機制：最多計算 24 個月
                if (count >= 24)
                    break;
            }
        }

        /// <summary>
        /// 取得該週的開始日（週一）
        /// </summary>
        private static DateTime GetWeekStart(DateTime date)
        {
            int diff = (7 + (date.DayOfWeek - DayOfWeek.Monday)) % 7;
            return date.AddDays(-diff).Date;
        }

        /// <summary>
        /// 取得該週指定星期幾的日期
        /// </summary>
        private static DateTime GetDateOfWeek(DateTime weekStart, DayOfWeek dayOfWeek)
        {
            int diff = dayOfWeek - DayOfWeek.Monday;
            if (diff < 0) diff += 7;
            return weekStart.AddDays(diff);
        }
    }

    /// <summary>
    /// 批次預約結果
    /// </summary>
    public class BatchBookingResult
    {
        public bool Success { get; set; }
        public List<BookingItem> SuccessfulBookings { get; set; }
        public List<FailedBookingItem> FailedBookings { get; set; }
        public string Message { get; set; }

        public BatchBookingResult()
        {
            SuccessfulBookings = new List<BookingItem>();
            FailedBookings = new List<FailedBookingItem>();
        }
    }

    /// <summary>
    /// 成功的預約項目
    /// </summary>
    public class BookingItem
    {
        public DateTime Date { get; set; }
        public string CaseId { get; set; }
        public string RoomId { get; set; }
    }

    /// <summary>
    /// 失敗的預約項目
    /// </summary>
    public class FailedBookingItem
    {
        public DateTime Date { get; set; }
        public string Reason { get; set; }
        public bool IsRoomOccupied { get; set; }
    }
}
