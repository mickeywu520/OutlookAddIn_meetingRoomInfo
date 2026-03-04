## POST - 新增會議室
http://192.168.0.13:100/api/MeetingRoom/addRent

## 帶入JSON
{"CaseId":"","RoomId":"R001","UserId":"11754","UserName":"","StartDate":"2026-03-03T09:30:00.000Z","EndDate":"2026-03-03T10:00:00.000Z","CreateTime":"2026-03-03T05:46:33.354Z","Subject":"軟體部會議","Remark":"磐儀#286","Cancel":false,"MeetingRoom":{"RoomId":"","Name":"","Type":"","Disable":false,"Remark":""}}

## 測試
{"CaseId":"","RoomId":"R003","UserId":"11754","UserName":"mickey[吳亞哲]","StartDate":"2026-03-04T09:30:00.000Z","EndDate":"2026-03-04T10:30:00.000Z","CreateTime":"2026-03-04T02:18:32.597Z","Subject":"[會議室預約] Test","Remark":"磐儀#11754","Cancel":false,"MeetingRoom":{"RoomId":"","Name":"","Type":"","Disable":false,"Remark":""}}

## RES
"1"
---

## POST - 獲取會議室最新租借情況
http://192.168.0.13:100/api/MeetingRoom/getRentRecord

## 帶入JSON
{"CaseId":"","RoomId":"","UserId":"","UserName":"","StartDate":"2026-02-26T00:00:00.000Z","EndDate":"2026-02-26T06:25:42.291Z","CreateTime":"2026-02-26T06:25:42.291Z","Subject":"","Remark":"","Cancel":false,"MeetingRoom":{"RoomId":"","Name":"","Type":"","Disable":false,"Remark":""}}

## RES
[
    {
        "UserName": "劉孟澤",
        "CaseId": "2026022601",
        "RoomId": "R001",
        "UserId": "11594",
        "StartDate": "2026-02-26T16:30:00",
        "EndDate": "2026-02-26T19:00:00",
        "CreateTime": "2025-03-27T16:11:16.573",
        "Subject": "MIS部門會議",
        "Remark": "磐儀#239",
        "Cancel": false
    },
    {
        "UserName": "蘇建宇",
        "CaseId": "R2026022608",
        "RoomId": "R005",
        "UserId": "11759",
        "StartDate": "2026-02-26T14:00:00",
        "EndDate": "2026-02-26T15:30:00",
        "CreateTime": "2026-02-11T14:20:06.153",
        "Subject": "面試",
        "Remark": "磐儀#277",
        "Cancel": false
    },
    {
        "UserName": "葉淑明",
        "CaseId": "R2026022609",
        "RoomId": "R002",
        "UserId": "11920",
        "StartDate": "2026-02-26T10:30:00",
        "EndDate": "2026-02-26T11:30:00",
        "CreateTime": "2026-02-11T15:19:50.147",
        "Subject": "TIOTA協會313研討會籌備",
        "Remark": "磐儀#373",
        "Cancel": false
    },
    {
        "UserName": "黃琪婷",
        "CaseId": "R2026022610",
        "RoomId": "R009",
        "UserId": "11955",
        "StartDate": "2026-02-26T09:00:00",
        "EndDate": "2026-02-26T18:00:00",
        "CreateTime": "2026-02-11T17:52:22.437",
        "Subject": "UL Audit",
        "Remark": "建康廠#680",
        "Cancel": false
    },
    {
        "UserName": "吳亞哲",
        "CaseId": "R2026022628",
        "RoomId": "R003",
        "UserId": "11754",
        "StartDate": "2026-02-26T18:00:00",
        "EndDate": "2026-02-26T18:30:00",
        "CreateTime": "2026-02-26T14:25:41.753",
        "Subject": "軟體部會議",
        "Remark": "磐儀#286",
        "Cancel": false
    }
]
---

## POST - 取消租借會議室
http://192.168.0.13:100/api/MeetingRoom/editRent

## 帶入JSON
{"UserName":"吳亞哲","CaseId":"R2026030436","RoomId":"R003","UserId":"11754","StartDate":"2026-03-04T09:30:00.000Z","EndDate":"2026-03-04T10:30:00.000Z","CreateTime":"2026-03-04T10:54:14.03","Subject":"軟體部會議","Remark":"磐儀#11754","Cancel":true}

## RES
"1"
---

## GET - 獲取會議室資訊
http://192.168.0.13:100/api/MeetingRoom/getroomlist

# RES JSON
[
    {
        "RoomId": "R001",
        "Name": "PARIS(原國際會議室)",
        "Type": "",
        "Sort": 1,
        "Remark": "財務部旁",
        "Disable": false
    },
    {
        "RoomId": "R002",
        "Name": "TAIPEI(原大會議室)",
        "Type": null,
        "Sort": 2,
        "Remark": "櫃檯後方大會議室",
        "Disable": false
    },
    {
        "RoomId": "R003",
        "Name": "SEOUL(首爾會議室)",
        "Type": null,
        "Sort": 3,
        "Remark": "首爾會議室、軟體部前面",
        "Disable": false
    },
    {
        "RoomId": "R005",
        "Name": "SAN JOSE(聖荷西會議室)",
        "Type": null,
        "Sort": 5,
        "Remark": "接待中心旁邊，5~6人",
        "Disable": false
    },
    {
        "RoomId": "R006",
        "Name": "LONDON(原業務會議室)",
        "Type": null,
        "Sort": 6,
        "Remark": "業務區(可容納8-10人)。為保護公司商業機密，僅供內部同仁使用，外賓一律謝絕\t",
        "Disable": false
    },
    {
        "RoomId": "R007",
        "Name": "Zoom",
        "Type": "虛擬",
        "Sort": 7,
        "Remark": "Zoom 視訊會議室",
        "Disable": false
    },
    {
        "RoomId": "R008",
        "Name": "建康廠-達文西",
        "Type": "健康廠",
        "Sort": 8,
        "Remark": "4~6人",
        "Disable": false
    },
    {
        "RoomId": "R009",
        "Name": "建康廠-拉菲爾",
        "Type": "健康廠",
        "Sort": 9,
        "Remark": "4~6人",
        "Disable": false
    },
    {
        "RoomId": "R010",
        "Name": "建康廠-米開朗基羅",
        "Type": "健康廠",
        "Sort": 10,
        "Remark": "大會議室，12~15人",
        "Disable": false
    }
]
---

## GET - 獲取使用者資訊
http://192.168.0.13:100/api/User/getAllUserListByEF

## RES
{
    "Code": "200",
    "Message": "取得資料成功",
    "Data": [
        {
            "UserId": "10001",
            "CompanyId": "Arbor",
            "UserNameZH": "李明",
            "UserNameEN": "Eric Lee",
            "ENName": "Eric Lee",
            "DepartmentId": "110R00",
            "DepartmentName": "董事長室",
            "DepartmentSort": 100,
            "JobTitleId": "001010",
            "Ext": "",
            "Email": "eric@arbor.com.tw"
        },
        {
            "UserId": "10550",
            "CompanyId": "Arbor",
            "UserNameZH": "陳榮昌",
            "UserNameEN": "Karl Chen",
            "ENName": "Karl Chen",
            "DepartmentId": "110R00",
            "DepartmentName": "董事長室",
            "DepartmentSort": 100,
            "JobTitleId": "003020",
            "Ext": "磐儀#306",
            "Email": "karl@arbor.com.tw"
        },
        {
            "UserId": "11621",
            "CompanyId": "Arbor",
            "UserNameZH": "李昇達",
            "UserNameEN": "Stanley Li",
            "ENName": "Stanley Li",
            "DepartmentId": "110R00",
            "DepartmentName": "董事長室",
            "DepartmentSort": 100,
            "JobTitleId": "003020",
            "Ext": "",
            "Email": "Stanley@arbor.com.tw"
        },
        {
            "UserId": "11754",
            "CompanyId": "Arbor",
            "UserNameZH": "吳亞哲",
            "UserNameEN": "Mickey Wu",
            "ENName": "Mickey Wu",
            "DepartmentId": "11D010",
            "DepartmentName": "軟體部",
            "DepartmentSort": 180,
            "JobTitleId": "005510",
            "Ext": "磐儀#286",
            "Email": "mickey@arbor.com.tw"
        },
        {
            "UserId": "11896",
            "CompanyId": "Arbor",
            "UserNameZH": "王文宏",
            "UserNameEN": "Kidd Wang",
            "ENName": "Kidd Wang",
            "DepartmentId": "11D900",
            "DepartmentName": "邊緣創新應用事業處研發部",
            "DepartmentSort": 110,
            "JobTitleId": "005020",
            "Ext": "磐儀#675",
            "Email": "kidd@arbor.com.tw"
        },
        {
            "UserId": "11841A",
            "CompanyId": "Arbor",
            "UserNameZH": "李宣萱",
            "UserNameEN": "Ariel",
            "ENName": "Ariel",
            "DepartmentId": "11T300",
            "DepartmentName": "經銷事業部",
            "DepartmentSort": 130,
            "JobTitleId": "005040",
            "Ext": "磐儀#317",
            "Email": "Ariel@arbor.com.tw"
        }
    ]
}
---