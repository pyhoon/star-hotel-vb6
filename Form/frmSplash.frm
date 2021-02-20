VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3930
      Left            =   120
      ScaleHeight     =   3930
      ScaleWidth      =   7560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   7560
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6720
         Top             =   960
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unlicensed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   555
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   6255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows XP or above"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   3960
         TabIndex        =   6
         Top             =   2520
         Width           =   3330
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   3960
         TabIndex        =   5
         Top             =   1560
         Width           =   3285
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: This software is full version. Please check license for more information."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   3540
         Width           =   6975
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   3180
         Width           =   6855
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   2940
         Width           =   6855
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loading... Please wait."
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         Top             =   1920
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 19/05/2018
' Modified On : 10/09/2018
' Modified On : 17/01/2019
' Modified On : 13/07/2019
Option Explicit
Private Const mstrModule As String = "Splash Screen"

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblCompanyProduct.Caption = COMPANY_PRODUCT_NAME
    lblProductName.Caption = App.ProductName
    lblCopyright.Caption = App.LegalCopyright '"Copyright 2014-2021" & App.LegalCopyright
    lblCompany.Caption = App.CompanyName
End Sub

' Version 1.1 to Version 1.2
' Copy old data to new database StarHotel.tmp.mdb and delete/backup old table
Private Sub Migrate_Database()
    Const mstrMethod As String = "Migrate Database"
    Dim rst As ADODB.Recordset
    Dim strLogFile As String
    Dim strText As String
On Error GoTo Exception
    OpenDB
    OpenData App.Path & "\Data\" & gstrCompanyID & ".tmp", gstrPassword
    
    ' ========================
    '  Version 1.1 to 1.2
    ' ========================
    strLogFile = App.Path & "\MigrateDB.log"
    ' WeeklyBooking - Not Copy
    ' UserGroup - Not Copy
    ' UserData
    Log2File strLogFile, "Migrate table [UserData]"
    Set rst = OpenTable("UserData")
    While Not rst.EOF
        SQL_UPDATE "UserData"
        SQL_SET_Long "UserGroup", rst("UserGroup").Value
        'SQL_SET_Text "UserID", rst("UserID").Value
        SQL_SET_Text "UserName", rst("UserName").Value
        SQL_SET_Text "UserPassword", rst("UserPassword").Value
        SQL_SET_Text "Salt", rst("Salt").Value
        SQL_SET_Long "LoginAttempts", rst("LoginAttempts").Value
        SQL_SET_Boolean "ChangePassword", rst("ChangePassword").Value
        SQL_SET_Boolean "Active", rst("Active").Value, False
        SQL_WHERE_Text "UserID", rst("UserID").Value
        QueryDataSQL gstrSQL
        rst.MoveNext
    Wend
    CloseRS rst
    Log2File strLogFile, "Migrate table [UserData] - success"
    ' RoomType
    Log2File strLogFile, "Migrate table [RoomType]"
    Set rst = OpenTable("RoomType")
    While Not rst.EOF
        SQL_UPDATE "RoomType"
        SQL_SET_Text "TypeShortName", rst("TypeShortName").Value
        If rst("TypeLongName").Value <> "" Then
            SQL_SET_Text "TypeLongName", rst("TypeLongName").Value
        Else
            SQL_SET_Text "TypeLongName", ""
        End If
        SQL_SET_Boolean "Active", rst("Active").Value, False
        SQL_WHERE_Long "ID", rst("ID").Value
        QueryDataSQL gstrSQL
        rst.MoveNext
    Wend
    CloseRS rst
    Log2File strLogFile, "Migrate table [RoomType] - success"
    ' Room
    Log2File strLogFile, "Migrate table [Room]"
    Set rst = OpenTable("Room")
    While Not rst.EOF
        SQL_UPDATE "Room"
        If rst("Maintenance").Value = True Then
            SQL_SET_Text "RoomStatus", "Maintenance"
            SQL_SET_Long "BookingID", 0
        Else
            If rst("BookingID").Value = 0 Then
                SQL_SET_Text "RoomStatus", "Open"
            Else
                If rst("xRoomStatus").Value = "Booked" Then
                    SQL_SET_Text "RoomStatus", "Booked"
                    SQL_SET_Long "BookingID", rst("BookingID").Value
                ElseIf rst("xRoomStatus").Value = "Occupied" Then
                    SQL_SET_Text "RoomStatus", "Occupied"
                    SQL_SET_Long "BookingID", rst("BookingID").Value
                Else
                    SQL_SET_Text "RoomStatus", "Open"
                    SQL_SET_Long "BookingID", 0
                End If
            End If
        End If
        SQL_SET_Text "RoomShortName", rst("RoomShortName").Value
        If rst("RoomLongName").Value <> "" Then
            SQL_SET_Text "RoomLongName", rst("RoomLongName").Value
        Else
            SQLText "RoomLongName = NULL"
        End If
        'SQL_SET_Text "RoomStatus", rst("xRoomStatus").Value
        SQL_SET_Text "RoomType", rst("RoomType").Value
        SQL_SET_Text "RoomLocation", rst("RoomLocation").Value
        SQL_SET_Double "RoomPrice", rst("RoomPrice").Value
        SQL_SET_Boolean "Breakfast", rst("Breakfast").Value
        If rst("BreakfastPrice").Value <> "" Then
            SQL_SET_Double "BreakfastPrice", rst("BreakfastPrice").Value
        Else
            SQLText "BreakfastPrice = 0"
        End If
        'SQL_SET_Double "BreakfastPrice", rst("BreakfastPrice").Value
        'SQL_SET_DateTime "CreatedDate", rst("CreatedDate").Value
        'SQL_SET_Text "CreatedBy", rst("CreatedBy").Value
        If rst("CreatedDate").Value <> "" Then
            SQL_SET_DateTime "CreatedDate", rst("CreatedDate").Value
        Else
            SQLText "CreatedDate = NULL"
        End If
        If rst("CreatedBy").Value <> "" Then
            SQL_SET_Text "CreatedBy", rst("CreatedBy").Value
        Else
            SQLText "CreatedBy = NULL"
        End If
        If rst("LastModifiedDate").Value <> "" Then
            SQL_SET_DateTime "LastModifiedDate", rst("LastModifiedDate").Value
        Else
            SQLText "LastModifiedDate = NULL"
        End If
        If rst("LastModifiedBy").Value <> "" Then
            SQL_SET_Text "LastModifiedBy", rst("LastModifiedBy").Value
        Else
            SQLText "LastModifiedBy = NULL"
        End If
        'SQL_SET_Boolean "Maintenance", rst("Maintenance").Value
        SQL_SET_Boolean "Active", rst("Active").Value, False
        SQL_WHERE_Integer "ID", rst("ID").Value
        QueryDataSQL gstrSQL
        rst.MoveNext
    Wend
    CloseRS rst
    Log2File strLogFile, "Migrate table [Room] - success"
    ' Report - Not Copy
    ' ModuleAccess
    Log2File strLogFile, "Migrate table [ModuleAccess]"
    Set rst = OpenTable("ModuleAccess")
    While Not rst.EOF
        Select Case rst("ModuleID").Value
        Case 1
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Booking"
        'SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 2
        Case 2
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Dashboard"
        'SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 1
        Case 3
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Maintain Room"
        'SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 9
        Case 4
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Maintain User"
        SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 10
        Case 5
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Access Control"
        SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 11
        Case 6
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Report"
        SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 4
        Case 7
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Edit Report"
        SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 5
        Case 8
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Edit Report (Expert)"
        SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 6
        Case 9
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Export Report"
        'SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 7
        Case 10
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Daily Booking Report"
        'SQL_SET_Text "ModuleType", "Report"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 12
        Case 11
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Weekly Booking Report"
        'SQL_SET_Text "ModuleType", "Report"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 13
        Case 12
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Monthly Booking Report"
        'SQL_SET_Text "ModuleType", "Report"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 14
        Case 13
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Weekly Booking Graph"
        'SQL_SET_Text "ModuleType", "Report"
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 15
        Case 14
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Shift Report for User"
        'SQL_SET_Text "ModuleType", "Report"
        SQL_SET_Boolean "Active", True
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 16
        Case 15
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Shift Report (All Users)"
        'SQL_SET_Text "ModuleType", "Report"
        SQL_SET_Boolean "Active", True
        SQL_SET_Boolean "Group1", rst("Group1").Value
        SQL_SET_Boolean "Group4", rst("Group4").Value, False
        SQL_WHERE_Integer "ModuleID", 17
'        Case 16
'        SQL_UPDATE "ModuleAccess"
'        SQL_SET_Text "ModuleDesc1", "Disabled Module" ' Void Booking
'        SQL_SET_Text "ModuleType", "Form"
'        SQL_SET_Boolean "Active", False
'        SQL_SET_Boolean "Group1", False
'        SQL_SET_Boolean "Group4", False, False
'        SQL_WHERE_Integer "ModuleID", 3
        Case 17
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Find Customer"
        SQL_SET_Text "ModuleType", "Form"
        SQL_SET_Boolean "Active", True
        SQL_SET_Boolean "Group1", True
        SQL_SET_Boolean "Group4", True, False
        SQL_WHERE_Integer "ModuleID", 8
        End Select
        QueryDataSQL gstrSQL
        rst.MoveNext
    Wend
    CloseRS rst
    Log2File strLogFile, "Migrate table [ModuleAccess] - success"
    ' LogRoom
    Log2File strLogFile, "Migrate table [LogRoom]"
    Set rst = OpenTable("LogRoom")
    While Not rst.EOF
        SQL_INSERT "LogRoom"
        SQLText "RoomID"
        SQLText "RoomShortName"
        SQLText "RoomLongName"
        SQLText "RoomStatus"
        SQLText "RoomType"
        SQLText "RoomLocation"
        SQLText "RoomPrice"
        SQLText "Breakfast"
        SQLText "BreakfastPrice"
        SQLText "CreatedDate"
        SQLText "CreatedBy"
        SQLText "LastModifiedDate"
        SQLText "LastModifiedBy"
        SQLText "Maintenance"
        SQLText "Active", False
        SQL_VALUES
        SQLData_Integer rst("RoomID").Value
        If rst("RoomShortName").Value <> "" Then
            SQLData_Text rst("RoomShortName").Value
        Else
            SQLText "NULL"
        End If
        If rst("RoomLongName").Value <> "" Then
            SQLData_Text rst("RoomLongName").Value
        Else
            SQLText "NULL"
        End If
        If rst("xRoomStatus").Value <> "" Then
            SQLData_Text rst("xRoomStatus").Value
        Else
            SQLText "NULL"
        End If
        If rst("RoomType").Value <> "" Then
            SQLData_Text rst("RoomType").Value
        Else
            SQLText "NULL"
        End If
        If rst("RoomLocation").Value <> "" Then
            SQLData_Text rst("RoomLocation").Value
        Else
            SQLText "NULL"
        End If
        If rst("RoomPrice").Value <> "" Then
            SQLData_Double rst("RoomPrice").Value
        Else
            SQLText "0"
        End If
        If rst("Breakfast").Value <> "" Then
            SQLData_Boolean rst("Breakfast").Value
        Else
            SQLText "FALSE"
        End If
        If rst("BreakfastPrice").Value <> "" Then
            SQLData_Double rst("BreakfastPrice").Value
        Else
            SQLText "0"
        End If
        If rst("CreatedDate").Value <> "" Then
            SQLData_DateTime rst("CreatedDate").Value
        Else
            SQLText "NULL"
        End If
        If rst("CreatedBy").Value <> "" Then
            SQLData_Text rst("CreatedBy").Value
        Else
            SQLText "NULL"
        End If
        If rst("LastModifiedDate").Value <> "" Then
            SQLData_DateTime rst("LastModifiedDate").Value
        Else
            SQLText "NULL"
        End If
        If rst("LastModifiedBy").Value <> "" Then
            SQLData_Text rst("LastModifiedBy").Value
        Else
            SQLText "NULL"
        End If
        SQLData_Boolean rst("Maintenance").Value
        SQLData_Boolean rst("Active").Value, False
        SQL_Close_Bracket
        QueryDataSQL gstrSQL
        rst.MoveNext
    Wend
    CloseRS rst
    Log2File strLogFile, "Migrate table [LogRoom] - success"

    ' LogBooking
    Log2File strLogFile, "Migrate table [LogBooking]"
    Set rst = OpenTable("LogBooking")
    While Not rst.EOF
        SQL_INSERT "LogBooking"
        SQLText "BookingID"
        SQLText "GuestName"
        SQLText "GuestPassport"
        SQLText "GuestOrigin"
        SQLText "GuestContact"
        SQLText "GuestEmergencyContactNo" ' <- GuestEmergencyContact
        SQLText "TotalGuest"
        SQLText "StayDuration"
        SQLText "BookingDate"
        SQLText "GuestCheckIN"
        SQLText "GuestCheckOUT"
        'SQLText "Status"
        SQLText "Remarks"
        SQLText "RoomID"
        SQLText "RoomNo"
        'SQLText "RoomStatus"
        SQLText "RoomType"
        SQLText "RoomLocation"
        SQLText "RoomPrice"
        SQLText "Breakfast"
        SQLText "BreakfastPrice"
        SQLText "SubTotal"
        SQLText "Deposit"
        SQLText "Payment"
        SQLText "Refund"
        SQLText "CreatedDate"
        SQLText "CreatedBy"
        SQLText "LastModifiedDate"
        SQLText "LastModifiedBy"
        SQLText "Active", False
        SQL_VALUES
        SQLData_Long rst("BookingID").Value, True, False
        SQLData_Text rst("GuestName").Value
        SQLData_Text rst("GuestPassport").Value
        SQLData_Text rst("GuestOrigin").Value
        SQLData_Text rst("GuestContact").Value
        SQLData_Text rst("GuestEmergencyContact").Value  ' -> GuestEmergencyContactNo
        SQLData_Integer rst("TotalGuest").Value
        SQLData_Integer rst("StayDuration").Value
        If rst("BookingDate").Value <> "" Then
            SQLData_DateTime rst("BookingDate").Value
        Else
            SQLText "NULL"
        End If
        If rst("GuestCheckIN").Value <> "" Then
            SQLData_DateTime rst("GuestCheckIN").Value
        Else
            SQLText "NULL"
        End If
        If rst("GuestCheckOUT").Value <> "" Then
            SQLData_DateTime rst("GuestCheckOUT").Value
        Else
            SQLText "NULL"
        End If
'        SQLData_DateTime rst("BookingDate").Value
'        SQLData_DateTime rst("GuestCheckIN").Value
'        SQLData_DateTime rst("GuestCheckOUT").Value
        'SQLData_Text rst("Status").Value
        SQLData_Text rst("Remarks").Value
        SQLData_Integer rst("RoomID").Value
        SQLData_Text rst("RoomNo").Value
        'If rst("xRoomStatus").Value <> "" Then
        '    SQLData_Text rst("xRoomStatus").Value
        'Else
        '    SQLText "NULL"
        'End If
        'SQLData_Text rst("RoomStatus").Value
        SQLData_Text rst("RoomType").Value
        SQLData_Text rst("RoomLocation").Value
        SQLData_Double rst("RoomPrice").Value
        SQLData_Boolean rst("Breakfast").Value
        SQLData_Double rst("BreakfastPrice").Value
        SQLData_Double rst("TotalDue").Value ' SubTotal
        SQLData_Double rst("PaidDeposit").Value
        SQLData_Double rst("PaidBalance").Value
        SQLData_Double 0
        If rst("CreatedDate").Value <> "" Then
            SQLData_DateTime rst("CreatedDate").Value
        Else
            SQLText "NULL"
        End If
        If rst("CreatedBy").Value <> "" Then
            SQLData_Text rst("CreatedBy").Value
        Else
            SQLText "NULL"
        End If
        If rst("LastModifiedDate").Value <> "" Then
            SQLData_DateTime rst("LastModifiedDate").Value
        Else
            SQLText "NULL"
        End If
        If rst("LastModifiedBy").Value <> "" Then
            SQLData_Text rst("LastModifiedBy").Value
        Else
            SQLText "NULL"
        End If
        'SQLData_DateTime rst("CreatedDate").Value
        'SQLData_Text rst("CreatedBy").Value
        'SQLData_DateTime rst("LastModifiedDate").Value
        'SQLData_Text rst("LastModifiedBy").Value
        SQLData_Boolean rst("Active").Value, False
        SQL_Close_Bracket
        QueryDataSQL gstrSQL
        rst.MoveNext
    Wend
    CloseRS rst
    Log2File strLogFile, "Migrate table [LogBooking] - success"
    ' Booking
    Log2File strLogFile, "Migrate table [Booking]"
    Set rst = OpenTable("Booking")
    While Not rst.EOF
        SQL_INSERT "Booking"
        SQLText "GuestName"
        SQLText "GuestPassport"
        SQLText "GuestOrigin"
        SQLText "GuestContact"
        SQLText "GuestEmergencyContactNo" ' <- GuestEmergencyContact
        SQLText "TotalGuest"
        SQLText "StayDuration"
        SQLText "BookingDate"
        SQLText "GuestCheckIN"
        SQLText "GuestCheckOUT"
        'SQLText "Status"
        SQLText "Remarks"
        SQLText "RoomID"
        SQLText "RoomNo"
        'SQLText "RoomStatus"
        SQLText "RoomType"
        SQLText "RoomLocation"
        SQLText "RoomPrice"
        SQLText "Breakfast"
        SQLText "BreakfastPrice"
        SQLText "SubTotal"
        SQLText "Deposit"
        SQLText "Payment"
        SQLText "Refund"
        SQLText "CreatedDate"
        SQLText "CreatedBy"
        SQLText "LastModifiedDate"
        SQLText "LastModifiedBy"
        SQLText "Active"
        SQLText "Temp", False
        SQL_VALUES
        SQLData_Text rst("GuestName").Value, True, False
        SQLData_Text rst("GuestPassport").Value
        SQLData_Text rst("GuestOrigin").Value
        SQLData_Text rst("GuestContact").Value
        SQLData_Text rst("GuestEmergencyContact").Value ' -> GuestEmergencyContactNo
        SQLData_Integer rst("TotalGuest").Value
        SQLData_Integer rst("StayDuration").Value
        SQLData_DateTime rst("BookingDate").Value
        SQLData_DateTime rst("GuestCheckIN").Value
        SQLData_DateTime rst("GuestCheckOUT").Value
        'SQLData_Text rst("Status").Value
        SQLData_Text rst("Remarks").Value
        SQLData_Integer rst("RoomID").Value
        SQLData_Text rst("RoomNo").Value
        'If rst("xRoomStatus").Value <> "" Then
        '    SQLData_Text rst("xRoomStatus").Value
        'Else
        '    SQLText "NULL"
        'End If
        'SQLData_Text rst("RoomStatus").Value
        SQLData_Text rst("RoomType").Value
        SQLData_Text rst("RoomLocation").Value
        SQLData_Double rst("RoomPrice").Value
        SQLData_Boolean rst("Breakfast").Value
        SQLData_Double rst("BreakfastPrice").Value
        SQLData_Double rst("TotalDue").Value ' SubTotal
        SQLData_Double rst("PaidDeposit").Value
        SQLData_Double rst("PaidBalance").Value
        SQLData_Double 0
        If rst("CreatedDate").Value <> "" Then
            SQLData_DateTime rst("CreatedDate").Value
        Else
            SQLText "NULL"
        End If
        If rst("CreatedBy").Value <> "" Then
            SQLData_Text rst("CreatedBy").Value
        Else
            SQLText "NULL"
        End If
        If rst("LastModifiedDate").Value <> "" Then
            SQLData_DateTime rst("LastModifiedDate").Value
        Else
            SQLText "NULL"
        End If
        If rst("LastModifiedBy").Value <> "" Then
            SQLData_Text rst("LastModifiedBy").Value
        Else
            SQLText "NULL"
        End If
        'SQLData_DateTime rst("CreatedDate").Value
        'SQLData_Text rst("CreatedBy").Value
        'SQLData_DateTime rst("LastModifiedDate").Value
        'SQLData_Text rst("LastModifiedBy").Value
        SQLData_Boolean rst("Active").Value
        SQLData_Boolean False, False
        SQL_Close_Bracket
        QueryDataSQL gstrSQL
        rst.MoveNext
    Wend
    CloseRS rst
    Log2File strLogFile, "Migrate table [Booking] - success"
    ' Company - Not Copy
    ' UserData
    Log2File strLogFile, "Update table [UserData]"
    ' Add Column DashboardBlink
    SQL_ALTER_TABLE "UserData"
    SQLText "ADD COLUMN Idle Integer", False
    QuerySQL gstrSQL
    Log2File strLogFile, "Update table [UserData] - success"
    CloseDB
    'MsgBox "Database update successfully. Please restart program.", vbExclamation, mstrMethod
    Exit Sub
Exception:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'Log2File strLogFile, mstrMethod, Err.Description
    WriteTextFile App.Path & "\Update_Error.log", Err.Description & vbCrLf & "SQL: " & gstrSQL
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

' Update Database if any new Tables/Columns added
Private Sub Update_Database(dblVersion As Double)
    Const mstrMethod As String = "Update Database"
    Dim rst As ADODB.Recordset
    Dim strText As String
On Error GoTo Exception
    OpenDB
    If dblVersion = 1.1 Then
        ' ========================
        '  Version 1.06 to 1.0.10
        ' ========================
        ' Add Column DashboardBlink
        SQL_ALTER_TABLE "UserData"
        SQLText "ADD COLUMN DashboardBlink Bit", False
        QuerySQL gstrSQL
        ' Rename Column TotalDue to SubTotal (Booking)
        SQL_ALTER_TABLE "Booking"
        SQLText "RENAME COLUMN TotalDue TO SubTotal", False
        QuerySQL gstrSQL
        ' Rename Column TotalDue to SubTotal (LogBooking)
        SQL_ALTER_TABLE "LogBooking"
        SQLText "RENAME COLUMN TotalDue TO SubTotal", False
        QuerySQL gstrSQL
        ' Update Database Version
        SQL_UPDATE "Company"
        SQL_SET_Double "DatabaseVersion", 1.1, False
        QuerySQL gstrSQL
    End If
    If dblVersion = 1.2 Then
        ' ========================
        '  Version 1.1 to 1.2
        ' ========================
        'Check Report ID = 5 already exist?
        SQL_SELECT_ALL "Report"
        SQL_WHERE_Integer "ReportID", 5
        Set rst = OpenRS(gstrSQL)
        If rst.EOF Then
            ' Add Report ID = 5
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
            SQLText "Active", False
            SQL_VALUES
            SQLData_Integer 5, True, False
            SQLData_Text "Shift Report for User"
            SQLData_Text "Shift Report for $UserID$"
            SQLData_Text " dated "
            SQLData_Boolean True
            SQLData_Text "CreatedDate"
            SQLData_Text "Single"
            SQLData_Text "Booking Report by User.rpt"
            strText = "SELECT CompanyName, StreetAddress, ContactNo, Format(ID,'100000') AS BookingID,"
            strText = strText & " RoomNo, RoomType, RoomPrice, SubTotal, Deposit, Payment, CreatedDate, CreatedBy"
            strText = strText & " FROM Company, Booking WHERE Booking.Active = TRUE AND CreatedBy = '$UserID$'"
            SQLData_Text strText
            SQLData_Text ""
            SQLData_Boolean True, False
            SQL_Close_Bracket
            QuerySQL gstrSQL
        End If
        CloseRS rst
        'Check Report ID = 6 already exist?
        SQL_SELECT_ALL "Report"
        SQL_WHERE_Integer "ReportID", 6
        Set rst = OpenRS(gstrSQL)
        If rst.EOF Then
            ' Add Report ID = 6
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
            SQLText "Active", False
            SQL_VALUES
            SQLData_Integer 6, True, False
            SQLData_Text "Shift Report (All Users)"
            SQLData_Text "Shift Report (Group by User ID)"
            SQLData_Text " dated "
            SQLData_Boolean True
            SQLData_Text "CreatedDate"
            SQLData_Text "Single"
            SQLData_Text "Booking Report by User.rpt"
            strText = "SELECT CompanyName, StreetAddress, ContactNo, Format(ID,'100000') AS BookingID,"
            strText = strText & " RoomNo, RoomType, RoomPrice, SubTotal, Deposit, Payment, CreatedDate, CreatedBy"
            strText = strText & " FROM Company, Booking WHERE Booking.Active = TRUE"
            SQLData_Text strText
            SQLData_Text ""
            SQLData_Boolean True, False
            SQL_Close_Bracket
            QuerySQL gstrSQL
        End If
        CloseRS rst
        ' Update Database Version
        SQL_UPDATE "Company"
        SQL_SET_Double "DatabaseVersion", 1.2, False
        QuerySQL gstrSQL
    End If
    If dblVersion = 1.3 Then
        'Check Report ID = 7 already exist?
        SQL_SELECT_ALL "Report"
        SQL_WHERE_Integer "ReportID", 7
        Set rst = OpenRS(gstrSQL)
        If rst.EOF Then
            ' Add Report ID = 7
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
            SQLData_Integer 7, True, False
            SQLData_Text "Official Receipt (Reprint)"
            SQLData_Text "OFFICIAL RECEIPT"
            SQLData_Text " "
            SQLData_Boolean False
            SQLData_Text "(None)"
            SQLData_Text "(None)"
            SQLData_Text "Official Receipt.rpt"
            strText = "SELECT C.CompanyName, C.StreetAddress, C.ContactNo, Format(B.ID, '100000') AS BookingID,"
            strText = strText & " B.GuestName, B.GuestCheckIN, B.GuestCheckOUT, B.RoomType,"
            strText = strText & " B.Payment, B.Refund, B.Payment-B.Refund AS Total, B.CreatedDate, '$UserID$' AS IssuedBy"
            strText = strText & " FROM Company C, Booking B WHERE Format(B.ID, '100000') = $BookingID$"
            SQLData_Text strText
            SQLData_Text ""
            strText = "SELECT CompanyName, StreetAddress, ContactNo, '000000' AS BookingID,"
            strText = strText & " '' AS GuestName, '2018-01-01' AS GuestCheckIN, '2018-01-01' AS GuestCheckOUT, '' AS RoomType,"
            strText = strText & "  0 AS Payment, 0 AS Refund, 0 AS Total, '' AS CreatedDate, '' AS IssuedBy"
            strText = strText & " FROM Company"
            SQLData_Text strText
            SQLData_Boolean True, False
            SQL_Close_Bracket
            QuerySQL gstrSQL
        End If
        CloseRS rst
        ' Enable New Access
        SQL_UPDATE "ModuleAccess"
        SQL_SET_Text "ModuleDesc1", "Official Receipt (Reprint)"
        SQL_SET_Text "ModuleType", "Report"
        SQL_SET_Boolean "Active", True
        SQL_SET_Boolean "Group1", True
        SQL_SET_Boolean "Group4", False, False
        SQL_WHERE_Integer "ModuleID", 18
        QuerySQL gstrSQL
        ' Update Database Version
        SQL_UPDATE "Company"
        SQL_SET_Double "DatabaseVersion", 1.3, False
        QuerySQL gstrSQL
    End If
    CloseDB
    'MsgBox "Database update successfully. Please restart program.", vbExclamation, mstrMethod
    Exit Sub
Exception:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'Log2File strLogFile, mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Timer1_Timer()
    Const mstrMethod As String = "Loading"
    Dim dblUserDBVersion As Double
    'Dim dblAppDBVersion As Double
    Dim strPath As String
    Dim strFile As String
On Error GoTo CheckErr
        DoEvents
        
        ' ========== ========== ========== Beta Version ========== ========== ==========
        Dim intDay As Integer
        intDay = DateDiff("d", "01 Jan 2021", Now)
        'If intDay < 0 Or intDay > 99 Then
        '    MsgBox "This system is expired!", vbOKOnly + vbExclamation, "System expired"
        '    'gstrPassword = ""
        '    End
        'End If
        
'        If intDay < 0 Or intDay > 364 Then
'            MsgBox "This system is expired on " & DateAdd("D", 364, "01 Jan 2021") & "!" & vbCrLf & _
'            "Please contact us at sales@computerise.my", vbOKOnly + vbExclamation, "System expired"
'            'gstrPassword = ""
'            End
'        End If
        
        ' ========== ========== ========== Beta Version ========== ========== ==========
        
        lblStatus.Caption = "Setting database password..."
        Me.Refresh
        gstrPassword = GenWord
    
        lblStatus.Caption = "Loading Configuration..."
        Me.Refresh

        If gstrDatabasePath = "" Then
            gstrCompanyID = "StarHotel"
            ' Database file extension
            gstrDatabaseExt = ".mdb"
            
            strPath = App.Path & "\Data\"
            strFile = gstrCompanyID & gstrDatabaseExt
            
            'Set Database path
            gstrDatabasePath = strPath & strFile
        End If
        
        lblStatus.Caption = "Locating database file..."
        Me.Refresh
        
        If Not FileExists(App.Path & "\Config.txt") Then
            Timer1.Enabled = False
            Unload Me
            With frmDatabase
                .txtFilePath.Text = strPath
                .txtFileName.Text = strFile
                .Show vbModal
            End With
            Exit Sub
        Else
            ReadTextFile "Config", 0, strPath
            ReadTextFile "Config", 1, strFile
            If strPath <> "" Then
                If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
                gstrDatabasePath = strPath
            End If
            If strFile <> "" Then
                gstrDatabasePath = strPath & strFile
            Else
                gstrDatabasePath = strPath & gstrCompanyID & gstrDatabaseExt
            End If
        End If
        
        If Not FileExists(gstrDatabasePath) Then
            Timer1.Enabled = False
            Unload Me
            With frmDatabase
                .txtFilePath.Text = strPath
                .txtFileName.Text = strFile
                .Show vbModal
            End With
            Exit Sub
        End If
        
        lblStatus.Caption = "Connecting to database..."
        Me.Refresh
        If ConnectDB = False Then
            Timer1.Enabled = False
            lblStatus.Caption = "Error during loading application."
            'lblStatus.Visible = True
            'MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
            'LogErrorText "Error", mstrMethod, Err.Description
            'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
            Unload Me
            Exit Sub
        End If
        
        lblStatus.Caption = "Checking Database Version..."
        Me.Refresh
        'dblAppDBVersion = 1.3
        dblUserDBVersion = DB_Version
        If dblUserDBVersion = 1.2 Then
            lblStatus.Caption = "Backing up database..."
            Me.Refresh
            If FileExists(strPath & strFile & ".bak") Then
                Kill strPath & strFile & ".bak"
            End If
            FileCopy strPath & strFile, strPath & strFile & ".bak"
            lblStatus.Caption = "Updating Database..."
            Me.Refresh
            Update_Database 1.3
            lblStatus.Caption = "Database has been updated."
            Me.Refresh
        ElseIf dblUserDBVersion = 1.1 Then
            lblStatus.Caption = "Migrating Database..."
            Me.Refresh
            Migrate_Database
            'lblStatus.Caption = "Database Migration completed."
            'Me.Refresh
            lblStatus.Caption = "Backing up database..."
            Me.Refresh
            If FileExists(App.Path & "\Data\" & gstrCompanyID & ".bak") Then
                Kill App.Path & "\Data\" & gstrCompanyID & ".bak"
            End If
            FileCopy App.Path & "\Data\" & gstrCompanyID & gstrDatabaseExt, App.Path & "\Data\" & gstrCompanyID & ".bak"
            Kill App.Path & "\Data\" & gstrCompanyID & gstrDatabaseExt
            lblStatus.Caption = "Copying new database..."
            Me.Refresh
            FileCopy App.Path & "\Data\" & gstrCompanyID & ".tmp", App.Path & "\Data\" & gstrCompanyID & gstrDatabaseExt
            'Kill App.Path & "\Data\" & gstrCompanyID & ".tmp" & gstrDatabaseExt
            Update_Database 1.2
            lblStatus.Caption = "Database has been updated."
            Me.Refresh
        ElseIf dblUserDBVersion < 1.1 Then
            lblStatus.Caption = "Updating Database..."
            Me.Refresh
            Update_Database 1.1
            'DB_Version_Update 1.1
            lblStatus.Caption = "Database has been updated."
            Me.Refresh
        Else
            lblStatus.Caption = "Database is already updated."
            Me.Refresh
        End If
        lblStatus.Caption = "Creating Reporting object..."
        Me.Refresh
        Set CrApplication = New CRAXDRT.Application
        lblStatus.Caption = "Loading completed."
        Me.Refresh
        lblStatus.Visible = False
'    End If
    Timer1.Enabled = False

    '''''Me.Hide
    frmUserLogin.Show
    Unload Me
    Exit Sub
CheckErr:
    lblStatus.Caption = "Error loading application."
    'lblStatus.Visible = True
    Timer1.Enabled = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
    Unload Me
End Sub
