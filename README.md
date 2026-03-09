# Outlook 會議室預約管理 Add-in

這是一個 Outlook VSTO Add-in 專案，用於整合公司內部會議室預約系統。當使用者在 Outlook 中建立或取消會議時，系統會自動同步預約或取消會議室。

## 功能特色

### 1. 自動會議室預約
- 當使用者發送會議邀請時，自動偵測 Location 欄位中的會議室名稱
- 自動呼叫 API 預約對應的會議室
- 支援多種會議室名稱識別（中英文、別名）

### 2. 自動取消會議室預約
- 當使用者取消會議時，自動偵測並取消對應的會議室預約
- 先查詢租借記錄取得 CaseId，再呼叫取消 API

### 3. 快速預約功能
- 透過 Ribbon 按開啟快速預約視窗
- 可視化顯示各會議室的可用時段（綠色=可預約、紅色=已占用、灰色=已逾時）
- 支援選擇不同日期查看預約狀況
- 點選「預約」按鈕後彈出確認對話框，可輸入會議主旨
- **立即預約機制**：確認後立即呼叫 API 預約會議室，避免被其他人搶走時段
- 預約成功後自動開啟 Outlook 新會議視窗，主旨和 Location 已自動填入

### 4. 會議室查詢功能
- 查詢今日/明日會議室預約狀況
- 自訂日期範圍查詢
- 匯出 CSV 報表

## 專案架構

```
OutlookAddIn_meetingRoomInfo/
├── ThisAddIn.cs              # Add-in 主要入口點
├── MeetingRoomRibbon.cs      # Ribbon UI 與功能邏輯
├── QuickBookingForm.cs       # 快速預約視窗
├── MeetingRoomResultForm.cs  # 查詢結果視窗
├── DateRangeForm.cs          # 日期選擇對話框
├── MeetingRecord.cs          # 資料模型類別
├── MeetingRoom.cs            # 會議室資料模型
└── packages.config           # NuGet 套件設定
```

## 核心元件說明

### ThisAddIn.cs
- **用途**: Add-in 的主要入口點，處理 Outlook 事件
- **主要功能**:
  - `Application_ItemSend`: 監聽會議發送事件，自動預約/取消會議室
  - `GetCurrentUserId()`: 從 Outlook 取得使用者員工編號 (Initials)
  - `GetCurrentUserExt()`: 從 API 取得使用者分機號碼
  - `BookMeetingRoomSync()`: 同步預約會議室
  - `CancelMeetingRoomSync()`: 同步取消會議室預約
  - `GetCaseIdFromRentRecord()`: 查詢租借記錄取得 CaseId

### MeetingRoomRibbon.cs
- **用途**: 定義 Ribbon UI 與使用者互動邏輯
- **主要功能**:
  - `btnQueryToday_Click`: 查詢今日預約
  - `btnQueryTomorrow_Click`: 查詢明日預約
  - `btnQueryRange_Click`: 自訂日期範圍查詢
  - `btnQuickBook_Click`: 開啟快速預約視窗
  - `FetchMeetingRooms()`: 取得會議室清單
  - `FetchMeetingRoomRecords()`: 取得預約記錄

### QuickBookingForm.cs
- **用途**: 快速預約介面
- **主要功能**:
  - 顯示會議室下拉選單
  - 日期選擇器（支援切換日期自動載入資料）
  - DataGridView 顯示各時段狀態（可預約/已占用）
  - 顯示預約人與會議主題
  - 多時段連續選擇預約

### MeetingRoomResultForm.cs
- **用途**: 查詢結果顯示介面
- **主要功能**:
  - 顯示會議室預約列表（含會議室名稱）
  - 雙擊查看詳細資訊
  - 匯出 CSV 功能

## 資料模型

### MeetingRecord
```csharp
public class MeetingRecord
{
    public string UserName { get; set; }
    public string RoomId { get; set; }
    public string StartDate { get; set; }
    public string EndDate { get; set; }
    public string Subject { get; set; }
    public string Remark { get; set; }
}
```

### RentRecord
```csharp
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
```

### MeetingRoom
```csharp
public class MeetingRoom
{
    public string RoomId { get; set; }
    public string Name { get; set; }
    public string Type { get; set; }
    public int Sort { get; set; }
    public string Remark { get; set; }
    public bool Disable { get; set; }
}
```

## API 整合

### 會議室相關 API

#### 1. 新增會議室預約
```
POST http://192.168.0.13:100/api/MeetingRoom/addRent
```

Request Body:
```json
{
  "CaseId": "",
  "RoomId": "R001",
  "UserId": "11754",
  "UserName": "吳亞哲",
  "StartDate": "2026-03-04T09:30:00.000Z",
  "EndDate": "2026-03-04T10:30:00.000Z",
  "CreateTime": "2026-03-04T02:18:32.597Z",
  "Subject": "軟體部會議",
  "Remark": "磐儀#286",
  "Cancel": false
}
```

#### 2. 取消會議室預約
```
POST http://192.168.0.13:100/api/MeetingRoom/editRent
```

Request Body:
```json
{
  "UserName": "吳亞哲",
  "CaseId": "R2026030436",
  "RoomId": "R003",
  "UserId": "11754",
  "StartDate": "2026-03-04T09:30:00.000Z",
  "EndDate": "2026-03-04T10:30:00.000Z",
  "Subject": "軟體部會議",
  "Remark": "磐儀#286",
  "Cancel": true
}
```

#### 3. 查詢租借記錄
```
POST http://192.168.0.13:100/api/MeetingRoom/getRentRecord
```

#### 4. 取得會議室清單
```
GET http://192.168.0.13:100/api/MeetingRoom/getroomlist
```

### 使用者相關 API

#### 取得使用者資訊
```
GET http://192.168.0.13:100/api/User/getAllUserListByEF
```

Response:
```json
{
  "Code": "200",
  "Message": "取得資料成功",
  "Data": [
    {
      "UserId": "11754",
      "UserNameZH": "吳亞哲",
      "Ext": "磐儀#286",
      "Email": "mickey@arbor.com.tw"
    }
  ]
}
```

## 會議室對應表

| 關鍵字 | RoomId | 完整名稱 |
|--------|--------|----------|
| PARIS / 國際會議室 | R001 | PARIS(原國際會議室) |
| TAIPEI / 大會議室 | R002 | TAIPEI(原大會議室) |
| SEOUL / 首爾 | R003 | SEOUL(首爾會議室) |
| SAN JOSE / 聖荷西 | R005 | SAN JOSE(聖荷西會議室) |
| LONDON / 業務會議室 | R006 | LONDON(原業務會議室) |
| ZOOM | R007 | Zoom |
| 達文西 | R008 | 建康廠-達文西 |
| 拉菲爾 | R009 | 建康廠-拉菲爾 |
| 米開朗基羅 | R010 | 建康廠-米開朗基羅 |

## 技術實作細節

### 1. 使用者身分識別
- 從 Outlook ExchangeUser 的 `Initials` 屬性取得員工編號
- Property Tag: `http://schemas.microsoft.com/mapi/proptag/0x3A0A001E`

### 2. 取消會議偵測
- 監聽 `Application.ItemSend` 事件
- 檢查 `AppointmentItem.MeetingStatus` 是否為 `olMeetingCanceled`

### 3. 分機號碼取得
- 呼叫使用者 API 取得 `Ext` 欄位
- 格式為 `"磐儀#XXX"` 或 `"建康廠#XXX"`
- 用於填入預約的 Remark 欄位

### 4. 跨執行緒處理
- 使用 `async/await` 進行非同步 API 呼叫
- 在事件處理器中使用 `.GetAwaiter().GetResult()` 進行同步等待

## 安裝需求

### 開發環境
- Visual Studio 2019 或更新版本
- .NET Framework 4.7.2 或更新版本
- Microsoft Office Outlook 2016 或更新版本

### 必要 NuGet 套件
- Newtonsoft.Json (JSON 序列化/反序列化)

## 部署方式

1. 在 Visual Studio 中建置專案
2. 發佈 ClickOnce 安裝程式，或
3. 手動複製 DLL 到使用者電腦並註冊

## 注意事項

1. **網路連線**: 需要連線至公司內網才能存取 API
2. **Outlook 權限**: 需要有 Exchange 帳號才能取得 Initials 欄位
3. **時區處理**: API 使用 UTC 時間，本地時間會自動轉換
4. **錯誤處理**: API 失敗時會顯示警告訊息，使用者可選擇是否繼續

## 版本歷史

### v1.0
- 初始版本
- 支援自動預約會議室
- 支援快速預約功能
- 支援查詢功能

### v1.1
- 新增自動取消會議室功能
- 修正 GetCurrentUserId 使用 Initials 欄位
- 新增分機號碼取得功能

### v1.2
- 快速預約支援切換日期自動載入資料
- 查詢結果顯示會議室完整名稱
- 快速預約新增「會議主題」欄位
- 改善 UI 版面配置適應不同解析度

### v1.3
- **快速預約流程優化**：
  - 點選「預約」後彈出確認對話框，可輸入會議主旨
  - **立即預約機制**：確認後立即呼叫 API 預約，避免時段被搶走
  - 預約成功後才開啟 Outlook 會議視窗
  - 移除發送會議時的重複預約檢查（已在確認時完成）
- **Location 欄位簡化**：只填入會議室名稱（不含 RoomId），避免主旨被截斷
- **已逾時時段顯示**：灰色標示已過期的時段，但仍顯示預約資訊

## 作者與維護

- 開發者: Mickey Wu
- 部門: MIS / 軟體部
- 分機: 磐儀#286
