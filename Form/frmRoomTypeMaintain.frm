VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoomTypeMaintain 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Type Maintenance"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   2280
   ClientWidth     =   15375
   Icon            =   "frmRoomTypeMaintain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   6240
      ScaleHeight     =   5145
      ScaleWidth      =   8865
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2280
      Width           =   8895
      Begin VB.TextBox txtTypeShortName 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   1
         ToolTipText     =   "Max length = 30"
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtTypeLongName 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   3480
         MaxLength       =   255
         TabIndex        =   2
         ToolTipText     =   "Max length = 255"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00303030&
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   3480
         TabIndex        =   3
         ToolTipText     =   "Enable/Disable this Room Type"
         Top             =   1680
         Value           =   1  'Checked
         Width           =   4815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   " Active"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " Short Description *"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " Long Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblRoomTypeID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   3480
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   240
      ScaleHeight     =   5145
      ScaleWidth      =   5865
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2280
      Width           =   5895
      Begin MSComctlLib.ListView lvRoomTypes 
         Height          =   4695
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         ForeColor       =   14737632
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Short Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Long Description"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Height          =   810
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   1429
      ButtonWidth     =   2064
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close (Esc)"
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Close (Esc)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear (Ctrl+C)"
            Key             =   "CLEAR"
            Object.ToolTipText     =   "Clear (Ctrl+C)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset (Ctrl+R)"
            Key             =   "RESET"
            Object.ToolTipText     =   "Reset (Ctrl+R)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save (Ctrl+S)"
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save Room Type (Ctrl+S)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "DELETE (Ctrl+D)"
            Key             =   "DELETE"
            Object.ToolTipText     =   "Delete Room Type (Ctrl+D)"
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmRoomTypeMaintain.frx":08CA
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   6
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   390
            Width           =   1620
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4800
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   -2147483633
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomTypeMaintain.frx":0BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomTypeMaintain.frx":14BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomTypeMaintain.frx":1DB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomTypeMaintain.frx":20E2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmRoomTypeMaintain.frx":2DBC
      Top             =   120
      Width           =   4050
   End
   Begin VB.Label lblBusinessName 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   10
      Top             =   240
      Width           =   15135
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
      TabIndex        =   9
      Top             =   900
      Width           =   5295
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
Attribute VB_Name = "frmRoomTypeMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
Option Explicit
Private Const mstrModule As String = "Room Type Maintenance"
Private Const COL_GRAY = &HE0E0E0
Private Const COL_PINK = &H8080FF
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
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    ListRoomType
    'lvRoomTypes.ListItems(1).Selected = True
    'lvRoomTypes_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyC And (Shift And vbCtrlMask) And tbrMenu.Buttons("CLEAR").Enabled Then
        ResetFields
    ElseIf KeyCode = vbKeyR And (Shift And vbCtrlMask) And tbrMenu.Buttons("RESET").Enabled Then
        lvRoomTypes_Click
    ElseIf KeyCode = vbKeyS And (Shift And vbCtrlMask) And tbrMenu.Buttons("SAVE").Enabled Then
        SaveRoomType
    'ElseIf KeyCode = vbKeyD And (Shift And vbCtrlMask) And tbrMenu.Buttons("DELETE").Enabled Then
    '    DeleteUser
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmRoomMaintain.Form_Load
    frmRoomMaintain.Show
End Sub

Private Sub lvRoomTypes_Click()
    If lvRoomTypes.ListItems.Count = 0 Then
        Exit Sub
    End If
    PopulateValues lvRoomTypes.SelectedItem.Text
End Sub

Private Sub lvRoomTypes_KeyUp(KeyCode As Integer, Shift As Integer)
    lvRoomTypes_Click
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "CLOSE"
            Unload Me
        Case "CLEAR"
            ResetFields
        Case "RESET"
            lvRoomTypes_Click
        Case "SAVE"
            SaveRoomType
        'Case "DELETE"
        '    DeleteUser
        Case Else 'Default
        
    End Select
End Sub

Private Sub ListRoomType()
    Const mstrMethod As String = "ListRoomType"
    Dim rst As ADODB.Recordset
    Dim List As ListItem
    Dim LSI As ListSubItem
    Dim i As Integer
On Error GoTo CheckErr
    lvRoomTypes.ListItems.Clear
    SQL_SELECT
    SQLText "ID"
    SQLText "TypeShortName"
    SQLText "TypeLongName"
    SQLText "Active", False
    SQL_FROM "RoomType"
    OpenDB
    Set rst = OpenRS(gstrSQL)
    While Not rst.EOF
        Set List = lvRoomTypes.ListItems.Add(, "i" & rst!ID, rst!ID, 0, 0)
        List.SubItems(1) = rst!TypeShortName
        If rst!TypeLongName <> "" Then
            List.SubItems(2) = rst!TypeLongName
        Else
            List.SubItems(2) = ""
        End If
        If rst!Active = True Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_GRAY
            Next
        Else
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_PINK
            Next
        End If
        lvRoomTypes.Refresh
        rst.MoveNext
    Wend
    CloseRS rst
    CloseDB
    lvRoomTypes_Click
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub PopulateValues(plngRoomTypeID As Long)
    Const mstrMethod As String = "PopulateValues"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "ID"
    SQLText "TypeShortName"
    SQLText "TypeLongName"
    SQLText "Active", False
    SQL_FROM "RoomType"
    SQL_WHERE_Long "ID", plngRoomTypeID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    ResetFields
    If Not rst.EOF Then
        lblRoomTypeID.Caption = rst("ID").Value
        txtTypeShortName.Text = rst("TypeShortName").Value
        If rst("TypeLongName").Value <> "" Then
            txtTypeLongName.Text = rst("TypeLongName").Value
        Else
            txtTypeLongName.Text = ""
        End If
        If rst("Active").Value = True Then
            chkActive.Value = vbChecked
        Else
            chkActive.Value = vbUnchecked
        End If
    Else
        MsgBox "Room Type not found", vbInformation, mstrMethod
    End If
    CloseRS rst
    CloseDB
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Sub SelectRoomType(plngRoomTypeID As Long)
    Const mstrMethod As String = "SelectRoom"
    Dim i As Integer
    Dim blnExist As Boolean
On Error GoTo CheckErr
    With lvRoomTypes
        For i = 1 To .ListItems.Count
            If .ListItems.Item(i).Text = plngRoomTypeID Then
                '.ListIndex = i
                .ListItems(i).Selected = True
                .SelectedItem.EnsureVisible
                blnExist = True
                lvRoomTypes_Click
                Exit For
            End If
        Next
        If blnExist = False Then
            lblRoomTypeID.Caption = plngRoomTypeID ' ""
            ResetFields
        End If
    End With
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

' Should not use this
'Private Sub DeleteRoomType()
'mstrMethod = "DeleteRoomType"
'Dim rst As ADODB.Recordset
'On Error GoTo CheckErr
'    If vbNo = MsgBox("Do you want to Delete this Room Type?", vbQuestion + vbYesNo, "Delete") Then
'        Exit Sub
'    End If
'    If cboUserGroup.ListIndex < 0 Then Exit Sub
'    SQL_SELECT_ALL "UserData"
'    SQL_WHERE_Text "UserID", Trim(txtUserID.Text)
'    OpenDB
'    Set rst = OpenSQL(gstrSQL)
'    If Not rst.EOF Then
'        SQL_DELETE "UserData"
'        SQL_WHERE_Text "UserID", Trim(txtUserID.Text)
'        QuerySQL gstrSQL
'        ResetFields
'        ListUsers
'    Else
'        MsgBox "User not found!", vbExclamation, mstrMethod
'    End If
'    CloseRS rst
'    CloseDB
'    Exit Sub
'CheckErr:
'    CloseRS rst
'    CloseDB
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
'    'LogErrorText "Error", mstrMethod, Err.Description
'    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
'End Sub

Private Sub SaveRoomType()
    Const mstrMethod As String = "Save Room Type"
    Dim rst As ADODB.Recordset
    Dim plngRoomTypeID As Long
    Dim mstrShortName As String
    Dim mstrLongName As String
On Error GoTo CheckErr
    plngRoomTypeID = ConvLng(lblRoomTypeID.Caption)
    mstrShortName = Trim(txtTypeShortName.Text)
    If mstrShortName = "" Then
        MsgBox "Please key in Room Type Short Description", vbExclamation, mstrMethod
        Exit Sub
    End If
    mstrLongName = Trim(txtTypeLongName.Text)
    If vbNo = MsgBox("Do you want to Save this Room Type?", vbQuestion + vbYesNo, mstrMethod) Then
        Exit Sub
    End If
    SQL_SELECT_ALL "RoomType"
    SQL_WHERE_Long "ID", plngRoomTypeID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        ' Check if any Room is this Type
        If chkActive.Value = vbUnchecked Then
            If IsRoomTypeUsed(mstrShortName) = True Then
                MsgBox "Room Type is currently used!", vbOKOnly + vbExclamation, "Room Type"
                CloseRS rst
                CloseDB
                Exit Sub
            End If
        End If
        SQL_UPDATE "RoomType"
        SQL_SET_Text "TypeShortName", mstrShortName
        SQL_SET_Text "TypeLongName", mstrLongName
        If chkActive.Value = vbChecked Then
            SQL_SET_Boolean "Active", True, False
        Else
            SQL_SET_Boolean "Active", False, False
        End If
        SQL_WHERE_Long "ID", plngRoomTypeID
        CloseRS rst
        QuerySQL gstrSQL
        MsgBox "Room Type successfully updated!", vbInformation, mstrMethod
    Else
        SQL_INSERT "RoomType"
        SQLText "TypeShortName"
        SQLText "TypeLongName"
        SQLText "Active", False
        SQL_VALUES
        SQLData_Text CheckInput(txtTypeShortName.Text)
        SQLData_Text CheckInput(txtTypeLongName.Text)
        If chkActive.Value = vbChecked Then
            SQLData_Boolean True, False
        Else
            SQLData_Boolean False, False
        End If
        SQL_Close_Bracket
        CloseRS rst
        QuerySQL gstrSQL
        MsgBox "New Room Type successfully added!", vbInformation, mstrMethod
    End If
    CloseDB
    ResetFields
    ListRoomType
    ' Recommend to log out and log in is required
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub ResetFields()
    lblRoomTypeID.Caption = "0"
    txtTypeShortName.Text = ""
    txtTypeLongName.Text = ""
    chkActive.Value = 0
End Sub

' Added on 06 May 2015
Private Function IsRoomTypeUsed(strRoomType As String) As Boolean
    Const mstrMethod As String = "IsRoomTypeUsed"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "COUNT(ID) AS RoomCount", False
    SQL_FROM "Room"
    SQL_WHERE_Text "RoomType", strRoomType
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("RoomCount") > 0 Then
            IsRoomTypeUsed = True
        Else
            IsRoomTypeUsed = False
        End If
    Else
        IsRoomTypeUsed = False
    End If
    Exit Function
CheckErr:
    CloseRS rst
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Function
