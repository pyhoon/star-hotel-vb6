VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   9645
   ClientLeft      =   6375
   ClientTop       =   3525
   ClientWidth     =   15375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   15105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   15135
      Begin MSComctlLib.ListView lvReports 
         Height          =   3615
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   -2147483633
         BackColor       =   5263440
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
            ImageIndex      =   1
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Report ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Report Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Report Title"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Report As On"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date Type"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date Field"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpTransDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   12582912
         CalendarTitleBackColor=   16744576
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   59310083
         CurrentDate     =   43450
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   " Report as at"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.Toolbar tbrButton 
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   1429
      ButtonWidth     =   2381
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close (Esc) "
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Close (Esc)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit (Ctrl+E) "
            Key             =   "EDIT"
            Object.ToolTipText     =   "Edit (Ctrl+E)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview (Ctrl+V) "
            Key             =   "PREVIEW"
            Object.ToolTipText     =   "Preview (Ctrl+V)"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmReport.frx":08CA
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   4
         Top             =   0
         Width           =   4935
         Begin VB.Timer tmrClock 
            Interval        =   1000
            Left            =   0
            Top             =   120
         End
         Begin VB.Label lblSystemDateTime 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tuesday, 30 September 2014 3:32 PM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   960
            TabIndex        =   6
            Top             =   120
            Width           =   3900
         End
         Begin VB.Label lblUserID 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "User ID : Admin"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   390
            Left            =   3240
            TabIndex        =   5
            Top             =   390
            Width           =   1620
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   -2147483633
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReport.frx":0BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReport.frx":14BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReport.frx":1D98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Computerise System Solutions 2014-2021. All rights reserved."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   9960
      TabIndex        =   8
      Top             =   900
      Width           =   5295
   End
   Begin VB.Label lblBusinessName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Room Booking System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   15135
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmReport.frx":2672
      Top             =   120
      Width           =   4050
   End
   Begin VB.Shape shpCopyright 
      BackColor       =   &H00505050&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00505050&
      FillStyle       =   0  'Solid
      Height          =   1200
      Left            =   0
      Top             =   0
      Width           =   15435
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
Option Explicit
Private Const mstrModule As String = "Report"
Private Const COL_GRAY = &HE0E0E0
Private Const COL_PINK = &H8080FF
Dim strDeveloperPassword As String
Dim intTick As Integer

Private Sub tmrClock_Timer()
    If Second(Now) = 0 Then
        lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    End If
    If gintUserIdle > 0 Then
        intTick = intTick + 1
        If intTick > gintUserIdle Then
            tmrClock.Enabled = False
            frmDialog.Show vbModal
        End If
    End If
End Sub

Private Sub Form_Activate()
    tmrClock.Enabled = True
    intTick = 0
End Sub

Private Sub Form_Deactivate()
    tmrClock.Enabled = False
End Sub

Private Sub Form_Load()
    Const mstrMethod As String = "Report Form_Load"
    strDeveloperPassword = "expert"
On Error GoTo CheckErr
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    'Me.Left = (mdiMain.Width - Me.Width) / 2
    'Me.Top = (mdiMain.Height - Me.Height) / 2 - 680
    dtpTransDate.Value = Date
    LoadReports
    'gblnReportSelect = True
    If UserAccessModule(MOD_REPORT_EDIT) = True Then
        tbrButton.Buttons("EDIT").Enabled = True
    Else
        tbrButton.Buttons("EDIT").Enabled = False
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const mstrMethod As String = "Report Form_Unload"
On Error GoTo CheckErr
    'gblnReportSelect = False
    frmDashboard.Show
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Const mstrMethod As String = "Report Form_KeyDown"
    Dim strSecret As String
On Error GoTo CheckErr
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyE And (Shift And vbCtrlMask) Then
        strSecret = InputBox("Please enter Developer password:", "Developer password", "")
        If strDeveloperPassword = strSecret Then
            If UserAccessModule(MOD_REPORT_EDIT) = True Then
                EditReport
            End If
        Else
            MsgBox "Developer password is invalid!", vbExclamation, "Access denied"
        End If
    ElseIf KeyCode = vbKeyV And (Shift And vbCtrlMask) Then
        If lvReports.SelectedItem.Text = "" Then
            MsgBox "Please select a Report to print.", vbInformation, mstrModule
            Exit Sub
        End If
        PrintReport lvReports.SelectedItem.Text
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub lvReports_DblClick()
    Const mstrMethod As String = "Report lvReports_DblClick"
On Error GoTo CheckErr
    PrintReport lvReports.SelectedItem.Text
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub tbrButton_ButtonClick(ByVal Button As MSComctlLib.Button)
    Const mstrMethod As String = "Report tbrButton_ButtonClick"
    Dim strSecret As String
On Error GoTo CheckErr
    Select Case Button.Key
    Case "CLOSE"
        Unload Me
    Case "EDIT"
        strSecret = InputBox("Please enter Developer password:", "Developer password", "")
        If strDeveloperPassword = strSecret Then
            If UserAccessModule(MOD_REPORT_EDIT) = True Then
                EditReport
            End If
        Else
            MsgBox "Developer password is invalid!", vbExclamation, "Access denied"
        End If
    Case "PREVIEW"
        If lvReports.SelectedItem.Text = "" Then
            MsgBox "Please select a Report to print.", vbInformation, mstrModule
            Exit Sub
        End If
        PrintReport lvReports.SelectedItem.Text
    End Select
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Sub LoadReports()
    Const mstrMethod As String = "LoadReports"
    Dim rst As ADODB.Recordset
    Dim List As ListItem
On Error GoTo CheckErr
    SQL_SELECT_ALL "Report"
    SQL_ORDER_BY "ReportID"
    OpenDB
    Set rst = OpenRS(gstrSQL)
    lvReports.ListItems.Clear
    Do While Not rst.EOF
        If UserAccessModule(rst!ReportID + 11) = True Then
            Set List = lvReports.ListItems.Add(, "i" & rst!ReportID, rst!ReportID, 0, 0)
            List.SubItems(1) = rst!ReportID
            List.SubItems(2) = rst!ReportName1
            List.SubItems(3) = rst!ReportTitle1
            List.SubItems(4) = rst!ReportAsOn1
            List.SubItems(5) = rst!DateType1
            List.SubItems(6) = rst!DateField1
            If rst!Active = True Then
                List.ForeColor = COL_GRAY
                List.ListSubItems(1).ForeColor = COL_GRAY
                List.ListSubItems(2).ForeColor = COL_GRAY
                List.ListSubItems(3).ForeColor = COL_GRAY
                List.ListSubItems(4).ForeColor = COL_GRAY
                List.ListSubItems(5).ForeColor = COL_GRAY
                List.ListSubItems(6).ForeColor = COL_GRAY
            Else
                List.ForeColor = COL_PINK
                List.ListSubItems(1).ForeColor = COL_PINK
                List.ListSubItems(2).ForeColor = COL_PINK
                List.ListSubItems(3).ForeColor = COL_PINK
                List.ListSubItems(4).ForeColor = COL_PINK
                List.ListSubItems(5).ForeColor = COL_PINK
                List.ListSubItems(6).ForeColor = COL_PINK
            End If
        End If
        rst.MoveNext
    Loop
    CloseRS rst
    CloseDB
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
    CloseRS rst
    CloseDB
End Sub

Public Sub EditReport()
    Dim intReportID As Integer
    Const mstrMethod As String = "EditReport"
On Error GoTo CheckErr
    If lvReports.SelectedItem.Text = "" Then
        MsgBox "Please select a Report to edit.", vbInformation, mstrMethod
        Exit Sub
    End If
    intReportID = lvReports.SelectedItem.ListSubItems(1)
    With frmReportMaintain
        .Show
        '.PopulateValues lvReports.SelectedItem.ListSubItems(1)
        .SelectReport intReportID
        .ZOrder 0
    End With
    Me.Hide
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Sub PrintReport(lngReportID As Long)
    Const mstrMethod As String = "PrintReport"
On Error GoTo CheckErr
    Dim rst As ADODB.Recordset
    Dim rep As New frmPrint
    Dim strDate As String
    Screen.MousePointer = vbHourglass
    SQL_SELECT_ALL "Report"
    SQL_WHERE_Long "ReportID", lngReportID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If Not rst!Active = True Then
            Screen.MousePointer = vbDefault
            MsgBox "This report has been disabled.", vbExclamation, mstrMethod
            CloseRS rst
            CloseDB
            Exit Sub
        Else
            gstrReportFileName = rst!ReportFile
            If FileExists(App.Path & "\Report\" & gstrReportFileName) = False Then
                Screen.MousePointer = vbDefault
                MsgBox "The report file is not found.", vbExclamation, mstrMethod
                CloseRS rst
                CloseDB
                Exit Sub
            End If
            gstrSQL = rst!ReportQuery
            If rst!ReportTitle1 <> "" Then
                gstrReportTitle = rst!ReportTitle1
                gstrReportTitle = Replace(gstrReportTitle, "$UserID$", gstrUserID)
            Else
                gstrReportTitle = ""
            End If
            
            If rst!DateField1 <> "(None)" Then
                If InStr(1, gstrSQL, "WHERE") > 0 Then
                    gstrSQL = gstrSQL & " AND "
                Else
                    gstrSQL = gstrSQL & " WHERE "
                End If
                If rst!DateField1 = "LastModifiedDate" Then
                    gstrSQL = gstrSQL & "LastModifiedDate"
                ElseIf rst!DateField1 = "CreatedDate" Then
                    gstrSQL = gstrSQL & "CreatedDate"
                ElseIf rst!DateField1 = "CreatedDate (All 7 days)" Then
                    gstrSQL = gstrSQL & "CreatedDate"
                    ' For Weekly Graph
                    UpdateWeekDayTable dtpTransDate.Value
                Else ' (None)
                    ' No date condition
                End If
                If rst!DateType1 = "Range" Then
                    gstrSQL = gstrSQL & " BETWEEN #" & FormatDate(dtpTransDate.Value) & " 12:00AM# AND #" & FormatDate(dtpTransDate.Value) & " 11:59:59PM#"
                    strDate = "from " & FormatDate(dtpTransDate.Value) & " to " & FormatDate(dtpTransDate.Value)
                ElseIf rst!DateType1 = "Weekly" Then
                    gstrSQL = gstrSQL & " BETWEEN #" & WeekDay1(dtpTransDate.Value) & " 12:00AM# AND #" & WeekDay7(dtpTransDate.Value) & " 11:59:59PM#"
                    strDate = "from " & WeekDay1(dtpTransDate.Value) & " to " & WeekDay7(dtpTransDate.Value)
                ElseIf rst!DateType1 = "Monthly" Then
                    gstrSQL = gstrSQL & " BETWEEN #" & MonthDay1(dtpTransDate.Value) & " 12:00AM# AND #" & MonthDay30(dtpTransDate.Value) & " 11:59:59PM#"
                    strDate = FormatMonthYear(dtpTransDate.Value)
                ElseIf rst!DateType1 = "Yearly" Then
                    gstrSQL = gstrSQL & " BETWEEN #" & YearDay1(dtpTransDate.Value) & " 12:00AM# AND #" & YearDay365(dtpTransDate.Value) & " 11:59:59PM#"
                    strDate = FormatYear(dtpTransDate.Value)
                ElseIf rst!DateType1 = "Since Start" Then
                    gstrSQL = gstrSQL & " BETWEEN #" & gdtmFiscalYearStart & " 12:00AM# AND #" & FormatDate(dtpTransDate.Value) & " 11:59:59PM#"
                    strDate = "from " & FormatDate(gdtmFiscalYearStart) & " to " & FormatDate(dtpTransDate.Value)
                ElseIf rst!DateType1 = "Single" Then ' Single
                    gstrSQL = gstrSQL & " BETWEEN #" & FormatDate(dtpTransDate.Value) & " 12:00AM# AND #" & FormatDate(dtpTransDate.Value) & " 11:59:59PM#"
                    strDate = FormatDate(dtpTransDate.Value)
                Else
                    ' Something else
                End If
            End If
                        
            ' Extend Query
            If rst!SubQuery <> "" Then
                gstrSQL = gstrSQL & " " & rst!SubQuery
            End If
            gstrSQL = Replace(gstrSQL, "$UserID$", gstrUserID)
            
            If InStr(1, gstrSQL, "$BookingID$") > 0 Then
                Dim lngBookingID As Long
                Dim strTemp As String
                strTemp = InputBox("Enter Booking No", "Booking No", 100001)
                lngBookingID = ConvLng(strTemp)
                If lngBookingID >= 0 Then
                    gstrSQL = Replace(gstrSQL, "$BookingID$", lngBookingID)
                End If
            End If
            'Debug.Print gstrSQL
            
            'Rep.Caption = rst!ReportName
            If rst!ShowReportAsOn = True Then
                If rst!ReportAsOn1 <> "" Then
                    gstrReportTitle = gstrReportTitle & rst!ReportAsOn1
                End If
                'gstrReportTitle = gstrReportTitle & FormatDate(dtpTransDate.Value)
                gstrReportTitle = gstrReportTitle & strDate
            End If
            ' Use alternate Query to show Company header
            If Not QueryHasData(gstrSQL) Then
                gstrSQL = rst!NullQuery
            End If
        End If
    End If
    'Debug.Print gstrSQL
    CloseRS rst
    CloseDB
    OpenDB
    Set rst = OpenRS(gstrSQL)
    
    'Show the Report
    rep.Show
    'MsgBox gstrSQL
    CloseRS rst
    CloseDB
    Screen.MousePointer = vbDefault
    Exit Sub
CheckErr:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
    CloseRS rst
    CloseDB
End Sub

Private Sub UpdateWeekDayTable(datReportDate As Date)
    Const mstrMethod As String = "UpdateWeekDayTable"
    Dim mstrSQL As String
    Dim i As Integer
    Screen.MousePointer = vbHourglass
On Error GoTo CheckErr
    OpenDB
    For i = 1 To 7
        mstrSQL = "UPDATE WeeklyBooking"
        mstrSQL = mstrSQL & " SET CreatedDate = #" & WeekDayN(datReportDate, i) & "#"
        mstrSQL = mstrSQL & " WHERE ID = " & CLng(i)
        QuerySQL mstrSQL
    Next
    CloseDB
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
    CloseDB
End Sub
