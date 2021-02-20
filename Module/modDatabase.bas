Attribute VB_Name = "modDatabase"
' Version      : 1.2.22
' Modified On  : 17/12/2014
' Descriptions : This Module is created by Aeric Poon
'                for all necessary database routine
' Updates:     : 1) Add function ConnectDB to test database connection
'              : 2) Add public variable DAT for connecting to old database
'              : 3) Add function OpenData ambigious to OpenDB
'              : 4) Add function CloseData ambigious to CloseDB
'              : 5) Add function OpenDataRS ambigious to OpenRS
'              : 6) Add function CloseDataRS ambigious to CloseRS
'              : 7) Add function OpenDataTable ambigious to OpenTable
'              : 8) Add function QueryDataSQL ambigious to QuerySQL

Option Explicit
Private Const mstrModule As String = "modDatabase"
Public ACN As ADODB.Connection
Public DAT As ADODB.Connection
Public gstrDatabaseExt As String
Public gstrDatabasePath As String
Public gstrPassword As String

Public Sub OpenDB()
Const mstrMethod As String = "OpenDB"
On Error GoTo CheckErr
    Set ACN = New ADODB.Connection
    ACN.Provider = "Microsoft.Jet.OLEDB.4.0"
    ACN.ConnectionString = "Data Source=" & gstrDatabasePath
    ACN.Properties("Jet OLEDB:Database Password") = gstrPassword
    ACN.Open
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Function ConnectDB() As Boolean
Const mstrMethod As String = "ConnectDB"
On Error GoTo CheckErr
    Set ACN = New ADODB.Connection
    ACN.Provider = "Microsoft.Jet.OLEDB.4.0"
    ACN.ConnectionString = "Data Source=" & gstrDatabasePath
    ACN.Properties("Jet OLEDB:Database Password") = gstrPassword
    ACN.Open
    ConnectDB = True
    Exit Function
CheckErr:
    ConnectDB = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Sub CloseDB()
Const mstrMethod As String = "CloseDB"
On Error GoTo CheckErr
    If ACN Is Nothing Then
    Else
        If ACN.State = adStateOpen Then
            ACN.Close
        End If
        Set ACN = Nothing
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Function CreateData() As Boolean
Const mstrMethod As String = "CreateData"
On Error GoTo CheckErr
    Dim cat As New ADOX.Catalog
    Dim strDBCon As String
    strDBCon = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDBCon = strDBCon & "Data Source=" & gstrDatabasePath & ";"
    strDBCon = strDBCon & "Jet OLEDB:Database Password=" & gstrPassword
    cat.Create strDBCon
    CreateData = True
    Exit Function
CheckErr:
    Set cat = Nothing
    CreateData = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function CreateDB() As Boolean
Const mstrMethod As String = "CreateDB"
On Error GoTo CheckErr
    OpenDB
    SQL_CREATE "Booking"
    SQL_COLUMN_ID
    SQL_COLUMN_TEXT "[GuestName]", 50
    SQL_COLUMN_TEXT "[GuestPassport]", 50
    SQL_COLUMN_TEXT "[GuestOrigin]", 50
    SQL_COLUMN_TEXT "[GuestContact]", 50
    SQL_COLUMN_TEXT "[GuestEmergencyContactName]", 50
    SQL_COLUMN_TEXT "[GuestEmergencyContactNo]", 50
    SQL_COLUMN_NUMBER "[TotalGuest]", "INTEGER", "0"
    SQL_COLUMN_NUMBER "[StayDuration]", "INTEGER", "0"
    SQL_COLUMN_DATETIME "[BookingDate]"
    SQL_COLUMN_DATETIME "[GuestCheckIN]"
    SQL_COLUMN_DATETIME "[GuestCheckOUT]"
    SQL_COLUMN_TEXT "[Remarks]"
    SQL_COLUMN_NUMBER "[RoomID]", "INTEGER", "0"
    SQL_COLUMN_TEXT "[RoomNo]", 50
    SQL_COLUMN_TEXT "[RoomType]", 50
    SQL_COLUMN_TEXT "[RoomLocation]", 50
    SQL_COLUMN_NUMBER "[RoomPrice]", "CURRENCY"
    SQL_COLUMN_YESNO "[Breakfast]"
    SQL_COLUMN_NUMBER "[BreakfastPrice]", "CURRENCY"
    SQL_COLUMN_NUMBER "[SubTotal]", "CURRENCY"
    SQL_COLUMN_NUMBER "[Deposit]", "CURRENCY"
    SQL_COLUMN_NUMBER "[Payment]", "CURRENCY"
    SQL_COLUMN_NUMBER "[Refund]", "CURRENCY"
    SQL_COLUMN_YESNO "[Active]", ""
    SQL_COLUMN_YESNO "[Temp]", "Yes"
    SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
    SQL_COLUMN_TEXT "[CreatedBy]", 50, ""
    SQL_COLUMN_DATETIME "[LastModifiedDate]"
    SQL_COLUMN_TEXT "[LastModifiedBy]", 50, "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "Company"
    SQL_COLUMN_TEXT "[CompanyName]"
    SQL_COLUMN_TEXT "[StreetAddress]"
    SQL_COLUMN_TEXT "[ContactNo]"
    SQL_COLUMN_DATETIME "[SystemStartDate]"
    SQL_COLUMN_TEXT "[ProductVersion]", 50
    SQL_COLUMN_NUMBER "[DatabaseVersion]", "DOUBLE", "0"
    SQL_COLUMN_TEXT "[CurrencySymbol]", 3
    SQL_COLUMN_TEXT "[F01]", 50
    SQL_COLUMN_YESNO "[Active]", "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "LogBooking"
    SQL_COLUMN_ID
    SQL_COLUMN_NUMBER "[BookingID]", "LONG", "0"
    SQL_COLUMN_TEXT "[GuestName]", 50
    SQL_COLUMN_TEXT "[GuestPassport]", 50
    SQL_COLUMN_TEXT "[GuestOrigin]", 50
    SQL_COLUMN_TEXT "[GuestContact]", 50
    SQL_COLUMN_TEXT "[GuestEmergencyContactName]", 50
    SQL_COLUMN_TEXT "[GuestEmergencyContactNo]", 50
    SQL_COLUMN_NUMBER "[TotalGuest]", "INTEGER", "0"
    SQL_COLUMN_NUMBER "[StayDuration]", "INTEGER", "0"
    SQL_COLUMN_DATETIME "[BookingDate]"
    SQL_COLUMN_DATETIME "[GuestCheckIN]"
    SQL_COLUMN_DATETIME "[GuestCheckOUT]"
    SQL_COLUMN_TEXT "[Remarks]"
    SQL_COLUMN_NUMBER "[RoomID]", "INTEGER", "0"
    SQL_COLUMN_TEXT "[RoomNo]", 50
    SQL_COLUMN_TEXT "[RoomType]", 50
    SQL_COLUMN_TEXT "[RoomLocation]", 50
    SQL_COLUMN_NUMBER "[RoomPrice]", "CURRENCY"
    SQL_COLUMN_YESNO "[Breakfast]"
    SQL_COLUMN_NUMBER "[BreakfastPrice]", "CURRENCY"
    SQL_COLUMN_NUMBER "[SubTotal]", "CURRENCY"
    SQL_COLUMN_NUMBER "[Deposit]", "CURRENCY"
    SQL_COLUMN_NUMBER "[Payment]", "CURRENCY"
    SQL_COLUMN_NUMBER "[Refund]", "CURRENCY"
    SQL_COLUMN_YESNO "[Active]", ""
    SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
    SQL_COLUMN_TEXT "[CreatedBy]", 50, ""
    SQL_COLUMN_DATETIME "[LastModifiedDate]"
    SQL_COLUMN_TEXT "[LastModifiedBy]", 50, "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "LogError"
    SQL_COLUMN_ID
    SQL_COLUMN_DATETIME "[LogDateTime]"
    SQL_COLUMN_TEXT "[LogErrorNum]", 50
    SQL_COLUMN_MEMO "[LogErrorDescription]"
    SQL_COLUMN_TEXT "[LogUserName]", 50
    SQL_COLUMN_TEXT "[LogModule]", 255
    SQL_COLUMN_TEXT "[LogMethod]", 255
    SQL_COLUMN_TEXT "[LogType]", 255, "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "LogRoom"
    SQL_COLUMN_ID
    SQL_COLUMN_NUMBER "[BookingID]", "LONG", "0"
    SQL_COLUMN_NUMBER "[RoomID]", "INTEGER", "0"
    SQL_COLUMN_TEXT "[RoomShortName]"
    SQL_COLUMN_TEXT "[RoomLongName]"
    SQL_COLUMN_TEXT "[RoomStatus]", 50
    SQL_COLUMN_TEXT "[RoomType]", 50
    SQL_COLUMN_TEXT "[RoomLocation]", 50
    SQL_COLUMN_NUMBER "[RoomPrice]", "CURRENCY"
    SQL_COLUMN_YESNO "[Breakfast]"
    SQL_COLUMN_NUMBER "[BreakfastPrice]", "CURRENCY"
    SQL_COLUMN_YESNO "[Maintenance]", ""
    SQL_COLUMN_YESNO "[Active]", ""
    SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
    SQL_COLUMN_TEXT "[CreatedBy]", 50, ""
    SQL_COLUMN_DATETIME "[LastModifiedDate]"
    SQL_COLUMN_TEXT "[LastModifiedBy]", 50, "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "ModuleAccess"
    SQL_COLUMN_NUMBER "[ModuleID]", "INTEGER", "0"
    SQL_COLUMN_TEXT "[ModuleDesc1]", 50
    SQL_COLUMN_TEXT "[ModuleType]", 50
    SQL_COLUMN_YESNO "[Group1]", ""
    SQL_COLUMN_YESNO "[Group2]", ""
    SQL_COLUMN_YESNO "[Group3]", ""
    SQL_COLUMN_YESNO "[Group4]", ""
    SQL_COLUMN_YESNO "[Active]", "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "Report"
    SQL_COLUMN_NUMBER "[ReportID]", "INTEGER", "0"
    SQL_COLUMN_TEXT "[ReportName1]", 255
    SQL_COLUMN_TEXT "[ReportTitle1]", 255
    SQL_COLUMN_TEXT "[ReportAsOn1]", 50
    SQL_COLUMN_YESNO "[ShowReportAsOn]", ""
    SQL_COLUMN_TEXT "[DateField1]", 50
    SQL_COLUMN_TEXT "[DateType1]", 50
    SQL_COLUMN_TEXT "[ReportFile]"
    SQL_COLUMN_MEMO "[ReportQuery]"
    SQL_COLUMN_TEXT "[SubQuery]"
    SQL_COLUMN_MEMO "[NullQuery]"
    SQL_COLUMN_YESNO "[Active]", "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "Room"
    SQL_COLUMN_ID "ID", True, False, True
    SQL_COLUMN_NUMBER "[BookingID]", "LONG", "0"
    SQL_COLUMN_TEXT "[RoomShortName]"
    SQL_COLUMN_TEXT "[RoomLongName]"
    SQL_COLUMN_TEXT "[RoomStatus]", 50
    SQL_COLUMN_TEXT "[RoomType]", 50
    SQL_COLUMN_TEXT "[RoomLocation]", 50
    SQL_COLUMN_NUMBER "[RoomPrice]", "CURRENCY"
    SQL_COLUMN_YESNO "[Breakfast]"
    SQL_COLUMN_NUMBER "[BreakfastPrice]", "CURRENCY"
    SQL_COLUMN_YESNO "[Maintenance]", ""
    SQL_COLUMN_YESNO "[Active]", ""
    SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
    SQL_COLUMN_TEXT "[CreatedBy]", 50, "System"
    SQL_COLUMN_DATETIME "[LastModifiedDate]"
    SQL_COLUMN_TEXT "[LastModifiedBy]", 50, "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "RoomType"
    SQL_COLUMN_ID
    SQL_COLUMN_TEXT "[TypeShortName]", 30
    SQL_COLUMN_TEXT "[TypeLongName]", 255
    SQL_COLUMN_YESNO "[Active]", "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "UserData"
    SQL_COLUMN_ID
    SQL_COLUMN_NUMBER "[UserGroup]", "LONG", "0"
    SQL_COLUMN_TEXT "[UserID]", 20
    SQL_COLUMN_TEXT "[UserName]", 50
    SQL_COLUMN_TEXT "[UserPassword]", 50
    SQL_COLUMN_TEXT "[Salt]", 50
    SQL_COLUMN_NUMBER "[Idle]", "INTEGER", "0"
    SQL_COLUMN_NUMBER "[LoginAttempts]", "LONG", "0"
    SQL_COLUMN_YESNO "[ChangePassword]"
    SQL_COLUMN_YESNO "[DashboardBlink]"
    SQL_COLUMN_YESNO "[Active]", "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "UserGroup"
    SQL_COLUMN_ID "GroupID"
    SQL_COLUMN_TEXT "[GroupName]", 20
    SQL_COLUMN_TEXT "[GroupDesc]", 255
    SQL_COLUMN_NUMBER "[SecurityLevel]", "LONG", "0"
    SQL_COLUMN_YESNO "[Active]", "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_CREATE "WeeklyBooking"
    SQL_COLUMN_NUMBER "[ID]", "LONG", "0"
    SQL_COLUMN_NUMBER "[RoomPrice]", "CURRENCY", "0"
    SQL_COLUMN_NUMBER "[BreakfastPrice]", "CURRENCY", "0"
    SQL_COLUMN_NUMBER "[SubTotal]", "CURRENCY", "0"
    SQL_COLUMN_NUMBER "[Deposit]", "CURRENCY", "0"
    SQL_COLUMN_NUMBER "[Payment]", "CURRENCY", "0"
    SQL_COLUMN_NUMBER "[Refund]", "CURRENCY", "0"
    SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
    SQL_COLUMN_TEXT "[CreatedBy]", 50, "", True, True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    CloseDB
    CreateDB = True
    Exit Function
CheckErr:
    CreateDB = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function CreateSampleData() As Boolean
Const mstrMethod As String = "CreateSampleData"
Dim mstrSalt As String
On Error GoTo CheckErr
    OpenDB
    ' Company
    SQL_INSERT "Company"
    SQLText "CompanyName"
    SQLText "StreetAddress"
    SQLText "ContactNo"
    SQLText "SystemStartDate"
    SQLText "DatabaseVersion"
    SQLText "CurrencySymbol"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Text "STAR HOTEL"
    SQLData_Text "9, Jalan Bintang, 50100 Kuala Lumpur, Malaysia"
    SQLData_Text "Tel/Fax : +603 - 4200 6336"
    SQLData_DateTime "1/1/2018"
    SQLData_Double 1.3
    SQLData_Text "MYR"
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    ' Module Access
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 1
    SQLData_Text "Dashboard"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 2
    SQLData_Text "Booking"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 3
    SQLData_Text "List Report"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 4
    SQLData_Text "Print Report"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 5
    SQLData_Text "Export Report"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 6
    SQLData_Text "Edit Report"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 7
    SQLData_Text "Edit Report (Expert)"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 8
    SQLData_Text "Find Customer"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 9
    SQLData_Text "Maintain Room"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 10
    SQLData_Text "Maintain User"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 11
    SQLData_Text "Access Control"
    SQLData_Text "Form"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 12
    SQLData_Text "Daily Booking Report"
    SQLData_Text "Report"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 13
    SQLData_Text "Weekly Booking Report"
    SQLData_Text "Report"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 14
    SQLData_Text "Monthly Booking Report"
    SQLData_Text "Report"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 15
    SQLData_Text "Weekly Booking Graph"
    SQLData_Text "Report"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 16
    SQLData_Text "Shift Report for User"
    SQLData_Text "Report"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 17
    SQLData_Text "Shift Report (All Users)"
    SQLData_Text "Report"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "ModuleAccess"
    SQLText "ModuleID"
    SQLText "ModuleDesc1"
    SQLText "ModuleType"
    SQLText "Group1"
    SQLText "Group2"
    SQLText "Group3"
    SQLText "Group4"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 18
    SQLData_Text "Official Receipt (Reprint)"
    SQLData_Text "Report"
    SQLData_Boolean True
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    ' Report
    SQL_INSERT "Report"
    SQLText "ReportID"
    SQLText "ReportName1"
    SQLText "ReportTitle1"
    SQLText "ReportAsOn1"
    SQLText "ShowReportAsOn"
    SQLText "DateField1"
    SQLText "DateType1"
    SQLText "ReportFile"
    SQLText "ReportQuery"
    SQLText "SubQuery"
    SQLText "NullQuery"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 1
    SQLData_Text "Daily Booking Report"
    SQLData_Text "Daily Booking"
    SQLData_Text " as on "
    SQLData_Boolean True
    SQLData_Text "CreatedDate"
    SQLData_Text "Single"
    SQLData_Text "Daily Booking Report.rpt"
    SQLData_Text "SELECT C.CompanyName, C.StreetAddress, C.ContactNo," & _
    " Format(B.ID,'100000') AS BookingID, B.BookingDate, B.GuestCheckIN," & _
    " B.GuestCheckOUT, B.GuestName, B.RoomNo, B.RoomType," & _
    " B.Deposit , B.Payment, B.CreatedDate, B.CreatedBy" & _
    " FROM Company C, Booking B" & _
    " WHERE B.Active = TRUE AND B.Temp = FALSE"
    SQLData_Text "ORDER BY B.ID"
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " Null AS BookingID, Null AS BookingDate, Null AS GuestCheckIN," & _
    " Null AS GuestCheckOUT, Null AS GuestName, Null AS RoomNo, Null AS RoomType," & _
    " 0 AS Deposit, 0 AS Payment, Null AS CreatedDate, Null AS CreatedBy" & _
    " FROM Company"
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "Report"
    SQLText "ReportID"
    SQLText "ReportName1"
    SQLText "ReportTitle1"
    SQLText "ReportAsOn1"
    SQLText "ShowReportAsOn"
    SQLText "DateField1"
    SQLText "DateType1"
    SQLText "ReportFile"
    SQLText "ReportQuery"
    SQLText "SubQuery"
    SQLText "NullQuery"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 2
    SQLData_Text "Weekly Booking Report"
    SQLData_Text "Weekly Booking"
    SQLData_Text " "
    SQLData_Boolean True
    SQLData_Text "CreatedDate"
    SQLData_Text "Weekly"
    SQLData_Text "Weekly Booking Report.rpt"
    SQLData_Text "SELECT C.CompanyName, C.StreetAddress, C.ContactNo," & _
    " Format(B.ID,'100000') AS BookingID, B.BookingDate, B.GuestCheckIN," & _
    " B.GuestCheckOUT, B.GuestName, B.RoomNo, B.RoomType," & _
    " B.Deposit , B.Payment, B.CreatedDate, B.CreatedBy" & _
    " FROM Company C, Booking B"
    SQLData_Text "ORDER BY B.ID"
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " Null AS BookingID, Null AS BookingDate, Null AS GuestCheckIN," & _
    " Null AS GuestCheckOUT, Null AS GuestName, Null AS RoomNo, Null AS RoomType," & _
    " 0 AS Deposit, 0 AS Payment, Null AS CreatedDate, Null AS CreatedBy" & _
    " FROM Company"
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
        
    SQL_INSERT "Report"
    SQLText "ReportID"
    SQLText "ReportName1"
    SQLText "ReportTitle1"
    SQLText "ReportAsOn1"
    SQLText "ShowReportAsOn"
    SQLText "DateField1"
    SQLText "DateType1"
    SQLText "ReportFile"
    SQLText "ReportQuery"
    SQLText "SubQuery"
    SQLText "NullQuery"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 3
    SQLData_Text "Monthly Booking Report"
    SQLData_Text "Monthly Booking"
    SQLData_Text " for "
    SQLData_Boolean True
    SQLData_Text "CreatedDate"
    SQLData_Text "Monthly"
    SQLData_Text "Monthly Booking Report.rpt"
    SQLData_Text "SELECT C.CompanyName, C.StreetAddress, C.ContactNo," & _
    " Format(B.ID,'100000') AS BookingID, B.BookingDate, B.GuestCheckIN," & _
    " B.GuestCheckOUT, B.GuestName, B.RoomNo, B.RoomType," & _
    " B.Deposit , B.Payment, B.CreatedDate, B.CreatedBy" & _
    " FROM Company C, Booking B"
    SQLData_Text "ORDER BY B.ID"
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " Null AS BookingID, Null AS BookingDate, Null AS GuestCheckIN," & _
    " Null AS GuestCheckOUT, Null AS GuestName, Null AS RoomNo, Null AS RoomType," & _
    " 0 AS Deposit, 0 AS Payment, Null AS CreatedDate, Null AS CreatedBy" & _
    " FROM Company"
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "Report"
    SQLText "ReportID"
    SQLText "ReportName1"
    SQLText "ReportTitle1"
    SQLText "ReportAsOn1"
    SQLText "ShowReportAsOn"
    SQLText "DateField1"
    SQLText "DateType1"
    SQLText "ReportFile"
    SQLText "ReportQuery"
    SQLText "SubQuery"
    SQLText "NullQuery"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 4
    SQLData_Text "Weekly Booking (Chart)"
    SQLData_Text "Weekly Booking"
    SQLData_Text " "
    SQLData_Boolean True
    SQLData_Text "CreatedDate (All 7 days)"
    SQLData_Text "Weekly"
    SQLData_Text "Weekly Booking Graph.rpt"
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " SUM(Payment) AS Total_Amount, DateValue(CreatedDate) AS ReportDate FROM Company," & _
    " ( SELECT Payment, CreatedDate FROM Booking" & _
    " WHERE Booking.Active = TRUE"
    SQLData_Text "UNION SELECT Payment, CreatedDate FROM WeeklyBooking)" & _
    " GROUP BY DateValue(CreatedDate), CompanyName, StreetAddress, ContactNo"
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " 0 AS Total_Amount, Null AS ReportDate" & _
    " FROM Company"
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "Report"
    SQLText "ReportID"
    SQLText "ReportName1"
    SQLText "ReportTitle1"
    SQLText "ReportAsOn1"
    SQLText "ShowReportAsOn"
    SQLText "DateField1"
    SQLText "DateType1"
    SQLText "ReportFile"
    SQLText "ReportQuery"
    SQLText "SubQuery"
    SQLText "NullQuery"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 5
    SQLData_Text "Shift Report by Staff"
    SQLData_Text "Booking Report by $UserID$"
    SQLData_Text " on "
    SQLData_Boolean True
    SQLData_Text "CreatedDate"
    SQLData_Text "Single"
    SQLData_Text "Booking Report.rpt"
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " BookingDate, Format(ID,'100000') AS BookingID," & _
    " GuestName, GuestCheckIN, GuestCheckOUT," & _
    " RoomNo, RoomType," & _
    " Deposit, Payment, Payment-Refund AS Total," & _
    " CreatedDate, CreatedBy" & _
    " FROM Company C, Booking B" & _
    " WHERE B.Active = YES And B.Temp = NO" & _
    " AND CreatedBy = '$UserID$'"
    SQLData_Text ""
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " Null AS BookingDate, Null AS BookingID," & _
    " Null AS GuestName, Null AS GuestCheckIN, Null AS GuestCheckOUT," & _
    " Null AS RoomNo, Null AS RoomType," & _
    " 0 AS Deposit, 0 AS Payment, 0 AS Total," & _
    " Null AS CreatedDate, Null AS CreatedBy" & _
    " FROM Company"
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "Report"
    SQLText "ReportID"
    SQLText "ReportName1"
    SQLText "ReportTitle1"
    SQLText "ReportAsOn1"
    SQLText "ShowReportAsOn"
    SQLText "DateField1"
    SQLText "DateType1"
    SQLText "ReportFile"
    SQLText "ReportQuery"
    SQLText "SubQuery"
    SQLText "NullQuery"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 6
    SQLData_Text "Shift Report by All Staff"
    SQLData_Text "Booking Report by All Staff"
    SQLData_Text " dated "
    SQLData_Boolean True
    SQLData_Text "CreatedDate"
    SQLData_Text "Single"
    SQLData_Text "Booking Report by Staff.rpt"
    SQLData_Text "SELECT C.CompanyName, C.StreetAddress, C.ContactNo," & _
    " Format(B.ID,'100000') AS BookingID, B.BookingDate," & _
    " B.GuestCheckIN, B.GuestCheckOUT," & _
    " B.GuestName, B.RoomNo, B.RoomType," & _
    " B.Deposit, B.Payment," & _
    " B.CreatedDate , B.CreatedBy" & _
    " FROM Company C, Booking B" & _
    " WHERE B.Active = YES AND B.Temp = NO"
    SQLData_Text ""
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " Null AS BookingID, Null AS BookingDate," & _
    " Null AS GuestCheckIN, Null AS GuestCheckOUT," & _
    " Null AS GuestName, Null AS RoomNo, Null AS RoomType," & _
    " 0 AS Deposit, 0 AS Payment," & _
    " Null AS CreatedDate, Null AS CreatedBy" & _
    " FROM Company"
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "Report"
    SQLText "ReportID"
    SQLText "ReportName1"
    SQLText "ReportTitle1"
    SQLText "ReportAsOn1"
    SQLText "ShowReportAsOn"
    SQLText "DateField1"
    SQLText "DateType1"
    SQLText "ReportFile"
    SQLText "ReportQuery"
    SQLText "SubQuery"
    SQLText "NullQuery"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 7
    SQLData_Text "Official Receipt (Reprint)"
    SQLData_Text "OFFICIAL RECEIPT"
    SQLData_Text " "
    SQLData_Boolean False
    SQLData_Text "(None)"
    SQLData_Text "(None)"
    SQLData_Text "Official Receipt.rpt"
    SQLData_Text "SELECT C.CompanyName, C.StreetAddress, C.ContactNo," & _
    " Format(B.ID, '100000') AS BookingID, B.GuestName, B.GuestCheckIN," & _
    " B.GuestCheckOUT, B.RoomType, B.Payment, B.Refund, B.Payment-B.Refund AS Total," & _
    " B.CreatedDate, '$UserID$' AS IssuedBy" & _
    " FROM Company C, Booking B" & _
    " WHERE Format(B.ID, '100000') = $BookingID$"
    SQLData_Text ""
    SQLData_Text "SELECT CompanyName, StreetAddress, ContactNo," & _
    " '000000' AS BookingID, '' AS GuestName, '2018-01-01' AS GuestCheckIN, '2018-01-01' AS GuestCheckOUT," & _
    " '' AS RoomType, 0 AS Payment, 0 AS Refund, 0 AS Total, '' AS CreatedDate, '' AS IssuedBy" & _
    " FROM Company"
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    ' UserGroup
    SQL_INSERT "UserGroup"
    SQLText "GroupID"
    SQLText "GroupName"
    SQLText "GroupDesc"
    SQLText "SecurityLevel"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 1
    SQLData_Text "Administrator"
    SQLData_Text "Highest Level User Group"
    SQLData_Long 99
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "UserGroup"
    SQLText "GroupID"
    SQLText "GroupName"
    SQLText "GroupDesc"
    SQLText "SecurityLevel"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 2
    SQLData_Text "Manager"
    SQLData_Text "Cannot access Admin level"
    SQLData_Long 98
    SQLData_Boolean False, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "UserGroup"
    SQLText "GroupID"
    SQLText "GroupName"
    SQLText "GroupDesc"
    SQLText "SecurityLevel"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 3
    SQLData_Text "Supervisor"
    SQLData_Text "Supervisor"
    SQLData_Long 20
    SQLData_Boolean False, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "UserGroup"
    SQLText "GroupID"
    SQLText "GroupName"
    SQLText "GroupDesc"
    SQLText "SecurityLevel"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Integer 4
    SQLData_Text "Clerk"
    SQLData_Text "Cashier"
    SQLData_Long 10
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    ' RoomType
    SQL_INSERT "RoomType"
    SQLText "TypeShortName"
    SQLText "TypeLongName"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Text "SINGLE BED ROOM"
    SQLData_Text ""
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "RoomType"
    SQLText "TypeShortName"
    SQLText "TypeLongName"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Text "DOUBLE BED ROOM"
    SQLData_Text ""
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "RoomType"
    SQLText "TypeShortName"
    SQLText "TypeLongName"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Text "TWIN BED ROOM"
    SQLData_Text ""
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    SQL_INSERT "RoomType"
    SQLText "TypeShortName"
    SQLText "TypeLongName"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Text "DORM"
    SQLData_Text ""
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    ' First Room
    SQL_INSERT "Room"
    SQLText "ID"
    SQLText "RoomShortName"
    SQLText "RoomLongName"
    SQLText "RoomStatus"
    SQLText "RoomType"
    SQLText "RoomLocation"
    SQLText "RoomPrice"
    SQLText "Breakfast"
    SQLText "BreakfastPrice"
    SQLText "Maintenance"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Long 1
    SQLData_Text "101"
    SQLData_Text ""
    SQLData_Text "Open"
    SQLData_Text "SINGLE BED ROOM"
    SQLData_Text "Level 1"
    SQLData_Double 100
    SQLData_Boolean True
    SQLData_Double 10
    SQLData_Boolean False
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
                
    'UserData
    mstrSalt = GenSalt(6)
    SQL_INSERT "UserData"
    SQLText "UserGroup"
    SQLText "UserID"
    SQLText "UserName"
    SQLText "UserPassword"
    SQLText "Salt"
    SQLText "Idle"
    SQLText "Loginattempts"
    SQLText "ChangePassword"
    SQLText "DashboardBlink"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Long 1
    SQLData_Text "Admin"
    SQLData_Text "Demo"
    SQLData_Text Encrypt("admin", mstrSalt)
    SQLData_Text mstrSalt
    SQLData_Integer 0
    SQLData_Integer 0
    SQLData_Boolean False
    SQLData_Boolean True
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    mstrSalt = GenSalt(6)
    SQL_INSERT "UserData"
    SQLText "UserGroup"
    SQLText "UserID"
    SQLText "UserName"
    SQLText "UserPassword"
    SQLText "Salt"
    SQLText "Idle"
    SQLText "Loginattempts"
    SQLText "ChangePassword"
    SQLText "DashboardBlink"
    SQLText "Active", False
    SQL_VALUES
    SQLData_Long 4
    SQLData_Text "Clerk"
    SQLData_Text "Receptionist"
    SQLData_Text Encrypt("clerk", mstrSalt)
    SQLData_Text mstrSalt
    SQLData_Integer 300
    SQLData_Integer 0
    SQLData_Boolean False
    SQLData_Boolean True
    SQLData_Boolean True, False
    SQL_Close_Bracket
    'Debug.Print gstrSQL
    QuerySQL gstrSQL
    
    CloseDB
    CreateSampleData = True
    Exit Function
CheckErr:
    CreateSampleData = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Function
    
Public Sub OpenData(mstrDatabasePath As String, mstrPassword As String)
Const mstrMethod As String = "OpenData"
On Error GoTo CheckErr
    Set DAT = New ADODB.Connection
    With DAT
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & mstrDatabasePath
        .Properties("Jet OLEDB:Database Password") = mstrPassword
        .Open
    End With
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
End Sub

Public Sub CloseData()
Const mstrMethod As String = "CloseData"
On Error GoTo CheckErr
    If DAT Is Nothing Then
    Else
        If DAT.State = adStateOpen Then
            DAT.Close
        End If
        Set DAT = Nothing
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
End Sub

Public Function ExecuteSelectSQL(pstrSQL As String, Optional aiCursorType As Integer = adOpenDynamic) As ADODB.Recordset
Const mstrMethod As String = "ExecuteSelectSQL"
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    If aiCursorType <> adOpenDynamic Then
        rst.Open pstrSQL, ACN, aiCursorType, adLockOptimistic
    Else
        rst.Open pstrSQL, ACN, adOpenForwardOnly, adLockOptimistic
    End If
    Set ExecuteSelectSQL = rst
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Sub UnlockDB(Con As ADODB.Connection)
Const mstrMethod As String = "UnlockDB"
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    OpenDB
    'gstrSQL = "SELECT * FROM NOTHING"
    gstrSQL = "SELECT * FROM UserData"
    rst.Open gstrSQL, Con, adOpenForwardOnly, adLockOptimistic, adCmdText
    rst.Close
    Set rst = Nothing
    CloseDB
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub
 
Public Function OpenSQL(pstrSQL As String) As ADODB.Recordset
Const mstrMethod As String = "OpenSQL"
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrSQL, ACN, adOpenForwardOnly, adLockOptimistic, adCmdText
    Set OpenSQL = rst
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function OpenQuery(pstrSQL As String) As ADODB.Recordset
Const mstrMethod As String = "OpenQuery"
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrSQL, ACN, adOpenForwardOnly, adLockOptimistic, adCmdText 'adCmdStoredProc
    Set OpenQuery = rst
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function OpenRS(pstrSQL As String) As ADODB.Recordset
Const mstrMethod As String = "OpenRS"
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrSQL, ACN, adOpenStatic, adLockPessimistic, adCmdText
    Set OpenRS = rst
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description & " SQL->" & pstrSQL
End Function

Public Sub CloseRS(rst As ADODB.Recordset)
Const mstrMethod As String = "CloseRS"
On Error GoTo CheckErr
    If rst Is Nothing Then
    Else
        If rst.State = adStateOpen Then
            rst.Close
        End If
        Set rst = Nothing
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Function OpenTable(pstrTable As String) As ADODB.Recordset
Const mstrMethod As String = "OpenTable"
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrTable, ACN, adOpenDynamic, adLockPessimistic, adCmdTable
    Set OpenTable = rst
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description & " Table->" & pstrTable
End Function

Public Sub QuerySQL(ByVal pstrSQL As String, Optional ByRef plngRecordsAffected As Long)
Const mstrMethod As String = "QuerySQL"
On Error GoTo CheckErr
    ACN.BeginTrans
    ACN.Execute pstrSQL, plngRecordsAffected
    ACN.CommitTrans
    Exit Sub
CheckErr:
    ACN.RollbackTrans
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description & " SQL->" & pstrSQL
End Sub

Public Function OpenDataRS(pstrSQL As String) As ADODB.Recordset
Const mstrMethod As String = "OpenDataRS"
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrSQL, DAT, adOpenStatic, adLockPessimistic, adCmdText
    Set OpenDataRS = rst
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
End Function

Public Sub CloseDataRS(rst As ADODB.Recordset)
Const mstrMethod As String = "CloseDataRS"
On Error GoTo CheckErr
    If rst Is Nothing Then
    Else
        If rst.State = adStateOpen Then
            rst.Close
        End If
        Set rst = Nothing
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
End Sub

Public Function OpenDataTable(pstrTable As String) As ADODB.Recordset
Const mstrMethod As String = "OpenDataTable"
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrTable, DAT, adOpenDynamic, adLockPessimistic, adCmdTable
    Set OpenDataTable = rst
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
End Function

Public Sub QueryDataSQL(ByVal pstrSQL As String, Optional ByRef plngRecordsAffected As Long)
Const mstrMethod As String = "QueryDataSQL"
On Error GoTo CheckErr
    DAT.BeginTrans
    DAT.Execute pstrSQL, plngRecordsAffected
    DAT.CommitTrans
    Exit Sub
CheckErr:
    DAT.RollbackTrans
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description & " SQL->" & pstrSQL
End Sub

'Public Function CompactDB(pstrOriginalFileName As String, pstrDestinationFileName As String) As Boolean
'mstrMethod = "CompactDB"
'Dim oJetEngine As New JRO.JetEngine
'Dim strSource As String
'Dim strDestination As String
'On Error GoTo CheckErr
'    strSource = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                "Data Source=" & pstrOriginalFileName & ";" & _
'                "Jet OLEDB:Database Password=" & gstrPassword & ";" & _
'                "Jet OLEDB:Engine Type=5;"
'
'    strDestination = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'              "Data Source=" & pstrDestinationFileName & ";" & _
'              "Jet OLEDB:Database Password=" & gstrPassword & ";" & _
'              "Jet OLEDB:Engine Type=5;"
'    'Set ACN = Nothing
'    'CrApplication.LogOffServer "P2smon.dll", "VMPC" ', "", ""
'    oJetEngine.CompactDatabase strSource, strDestination
'    Set oJetEngine = Nothing
'    CompactDB = True
'    Exit Function
'CheckErr:
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
'    'LogError "Error", mstrMethod, Err.Description
'    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
'    Set oJetEngine = Nothing
'    CompactDB = False
'End Function
