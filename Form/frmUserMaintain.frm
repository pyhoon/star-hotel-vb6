VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserMaintain 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Maintenance"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   2280
   ClientWidth     =   15375
   Icon            =   "frmUserMaintain.frx":0000
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
      Height          =   5655
      Left            =   6240
      ScaleHeight     =   5625
      ScaleWidth      =   8865
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2280
      Width           =   8895
      Begin VB.TextBox txtUserID 
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
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "User ID"
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtName 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "User Name"
         Top             =   1080
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
         Left            =   3600
         TabIndex        =   10
         ToolTipText     =   "Enable/Disable this user."
         Top             =   2040
         Value           =   1  'Checked
         Width           =   4815
      End
      Begin VB.CheckBox chkChange 
         BackColor       =   &H00303030&
         Caption         =   "Change Password when login"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   3600
         TabIndex        =   12
         ToolTipText     =   "When checked, this user needs to Change Password when login."
         Top             =   3000
         Width           =   4815
      End
      Begin VB.CheckBox chkReset 
         BackColor       =   &H00303030&
         Caption         =   "Reset login attempts (0)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   3600
         TabIndex        =   11
         ToolTipText     =   "Reset freezed user."
         Top             =   2520
         Width           =   4815
      End
      Begin VB.CheckBox chkUpdatePassword 
         BackColor       =   &H00303030&
         Caption         =   "Update Password now"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   360
         Left            =   3600
         TabIndex        =   13
         ToolTipText     =   "Update Password for this user."
         Top             =   3480
         Width           =   4815
      End
      Begin VB.ComboBox cboUserGroup 
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
         ItemData        =   "frmUserMaintain.frx":0EB2
         Left            =   3600
         List            =   "frmUserMaintain.frx":0EB4
         MousePointer    =   1  'Arrow
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Display the user group"
         Top             =   120
         Width           =   4815
      End
      Begin VB.TextBox txtPasswordOld 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   8
         PasswordChar    =   "•"
         TabIndex        =   15
         ToolTipText     =   "Current password"
         Top             =   3960
         Width           =   4815
      End
      Begin VB.TextBox txtPasswordNew 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   8
         PasswordChar    =   "•"
         TabIndex        =   17
         ToolTipText     =   "Password to change"
         Top             =   4440
         Width           =   4815
      End
      Begin VB.TextBox txtPasswordConfirm 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   8
         PasswordChar    =   "•"
         TabIndex        =   19
         ToolTipText     =   "Confirm new password"
         Top             =   4920
         Width           =   4815
      End
      Begin VB.TextBox txtIdle 
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         Enabled         =   0   'False
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
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   8
         ToolTipText     =   "Value between 0 to 3600"
         Top             =   1560
         Width           =   4815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00303030&
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
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " User ID"
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
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " User Name"
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
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " Old Password"
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
         Left            =   360
         TabIndex        =   14
         Top             =   3960
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " New Password"
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
         Left            =   360
         TabIndex        =   16
         Top             =   4440
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " Confirm Password"
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
         Left            =   360
         TabIndex        =   18
         Top             =   4920
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " User Group"
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
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " Idle Time (seconds)"
         Enabled         =   0   'False
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
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Auto log-out when idle (secs)"
         Top             =   1560
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   240
      ScaleHeight     =   5625
      ScaleWidth      =   5865
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2280
      Width           =   5895
      Begin MSComctlLib.ListView lvUsers 
         Height          =   5175
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   9128
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
            Text            =   "User ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "User Name"
            Object.Width           =   7070
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Height          =   810
      Left            =   0
      TabIndex        =   0
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
            Object.ToolTipText     =   "Save User (Ctrl+S)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "DELETE (Ctrl+D)"
            Key             =   "DELETE"
            Object.ToolTipText     =   "Delete User (Ctrl+D)"
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmUserMaintain.frx":0EB6
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   21
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   390
            Width           =   1620
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5880
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
               Picture         =   "frmUserMaintain.frx":11D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserMaintain.frx":1AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserMaintain.frx":239C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUserMaintain.frx":26CE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmUserMaintain.frx":33A8
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
      TabIndex        =   25
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
      TabIndex        =   24
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
Attribute VB_Name = "frmUserMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
Option Explicit
Private Const mstrModule As String = "User Maintenance"
Private Const COL_PINK = &H8080FF
Private Const COL_GREEN = &H76E600
Private Const COL_GRAY = &HE0E0E0
Private Const COL_DISABLED = &H505050
Private Const COL_ENABLED = &HC0FFFF
Dim intTick As Integer

Private Sub cboUserGroup_Click()
    If cboUserGroup.ListIndex > 0 Then
        txtIdle.Enabled = True
    Else
        txtIdle.Enabled = False ' Any point?
    End If
End Sub

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
    Const mstrMethod As String = "Form_Load"
On Error GoTo CheckErr
    'Left = (mdiMain.ScaleWidth - Width) / 2
    'Top = (mdiMain.ScaleHeight - Height) / 2
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    PopulateGroupName
    'PopulateThemeName
    ListUsers
    'lvUsers.ListItems(1).Selected = True
    'lvUsers_Click
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrModule 'mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyC And (Shift And vbCtrlMask) And tbrMenu.Buttons("CLEAR").Enabled Then
        ResetFields
    ElseIf KeyCode = vbKeyR And (Shift And vbCtrlMask) And tbrMenu.Buttons("RESET").Enabled Then
        lvUsers_Click
    ElseIf KeyCode = vbKeyS And (Shift And vbCtrlMask) And tbrMenu.Buttons("SAVE").Enabled Then
        SaveUser
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const mstrMethod As String = "Form_Unload"
On Error GoTo CheckErr
    frmDashboard.Show
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrModule 'mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub lvUsers_Click()
    Const mstrMethod As String = "lvUsers_Click"
On Error GoTo CheckErr
    If lvUsers.ListItems.Count = 0 Then
        Exit Sub
    End If
    PopulateValues lvUsers.SelectedItem.Text
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrModule 'mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub lvUsers_KeyUp(KeyCode As Integer, Shift As Integer)
    Const mstrMethod As String = "lvUsers_KeyUp"
On Error GoTo CheckErr
    lvUsers_Click
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrModule 'mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "CLOSE"
            Unload Me
        Case "CLEAR"
            ResetFields
        Case "RESET"
            lvUsers_Click
        Case "SAVE"
            SaveUser
    '    Case "DELETE"
    '        DeleteUser
        Case Else 'Default
            Exit Sub
    End Select
End Sub

Private Sub chkUpdatePassword_Click()
    If chkUpdatePassword.Value = vbChecked Then
        txtPasswordOld.Enabled = True
        txtPasswordNew.Enabled = True
        txtPasswordConfirm.Enabled = True
        txtPasswordOld.BackColor = COL_ENABLED
        txtPasswordNew.BackColor = COL_ENABLED
        txtPasswordConfirm.BackColor = COL_ENABLED
        txtPasswordOld.SetFocus
    Else
        txtPasswordOld.Enabled = False
        txtPasswordNew.Enabled = False
        txtPasswordConfirm.Enabled = False
        txtPasswordOld.BackColor = COL_DISABLED
        txtPasswordNew.BackColor = COL_DISABLED
        txtPasswordConfirm.BackColor = COL_DISABLED
    End If
End Sub

Private Sub PopulateGroupName()
    Const mstrMethod As String = "PopulateGroupName"
    Dim rst As ADODB.Recordset
    Dim i As Integer
On Error GoTo CheckErr
    OpenDB
    Set rst = GetList("UserGroup", "GroupID", "GroupName", "Active = TRUE", "SecurityLevel", False)
    cboUserGroup.Clear
    While Not rst.EOF
        cboUserGroup.AddItem rst("GroupName").Value
        cboUserGroup.ItemData(i) = rst("GroupID").Value
        i = i + 1
        rst.MoveNext
    Wend
    CloseRS rst
    CloseDB
    Exit Sub
CheckErr:
    CloseRS rst
    'CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub ListUsers()
    Const mstrMethod As String = "ListUsers"
    Dim rst As ADODB.Recordset
    Dim List As ListItem
    Dim LSI As ListSubItem
    Dim i As Integer
On Error GoTo CheckErr
    lvUsers.ListItems.Clear
    SQL_SELECT
    SQLText "ID"
    SQLText "UserID"
    SQLText "UserName"
    SQLText "UserGroup"
    SQLText "Active", False
    SQL_FROM "UserData"
    OpenDB
    Set rst = OpenRS(gstrSQL)
    While Not rst.EOF
        Set List = lvUsers.ListItems.Add(, "i" & rst!ID, rst!ID, 0, 0)
        List.SubItems(1) = rst!UserID
        List.SubItems(2) = rst!UserName
        If rst!Active = False Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_PINK
            Next
        ElseIf rst!UserGroup = 1 Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_GREEN
            Next
        Else
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_GRAY
            Next
        End If
        lvUsers.Refresh
        rst.MoveNext
    Wend
    CloseRS rst
    CloseDB
    lvUsers_Click
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub PopulateValues(lngUserID As Long)
    Const mstrMethod As String = "PopulateValues"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    'gstrSQL = "SELECT GroupName, ID, UserID, UserName, Active, LoginAttempts, ChangePassword" & _
            " FROM UserData LEFT JOIN UserGroup ON UserData.UserGroup = UserGroup.GroupID" & _
            " WHERE UserData.ID = " & lngUserID
    SQL_SELECT
    SQLText "G.GroupName"
    SQLText "D.ID"
    SQLText "D.UserID"
    SQLText "D.UserName"
    'SQLText "Theme.ThemeName"
    SQLText "D.Active"
    SQLText "D.Idle"
    SQLText "D.LoginAttempts"
    SQLText "D.ChangePassword", False
    SQL_FROM "UserData", "D"
    SQL_INNER_JOIN "UserGroup", "G"
    SQL_ON "D", "UserGroup", "G", "GroupID"
    SQL_WHERE_Long "D.ID", lngUserID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    ResetFields
    If Not rst.EOF Then
        If rst("UserID").Value <> "" Then
            txtUserID.Text = Trim(rst("UserID").Value)
        Else
            txtUserID.Text = ""
        End If
        If rst("UserName").Value <> "" Then
            txtName.Text = Trim(rst("UserName").Value)
        Else
            txtName.Text = ""
        End If
        cboUserGroup.Text = rst("GroupName").Value
        If rst("Idle").Value <> "" Then
            txtIdle.Text = ConvInt(rst("Idle").Value)
        Else
            txtIdle.Text = "0"
        End If
        If rst("Active").Value = True Then
            chkActive.Value = vbChecked
        Else
            chkActive.Value = vbUnchecked
        End If
        If rst("ChangePassword").Value = True Then
            chkChange.Value = vbChecked
        Else
            chkChange.Value = vbUnchecked
        End If
        chkReset.Caption = "Reset login attempts (" & rst("LoginAttempts").Value & ")"
    Else
        MsgBox "User not found", vbInformation, mstrMethod
    End If
    'txtName.SelStart = Len(txtName.Text)
    'txtName.SetFocus
'    tbrMenu.Buttons("DELETE").Enabled = True
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

' Should not use this
'Private Sub DeleteUser()
'   mstrMethod = "DeleteUser"
'   Dim rst As ADODB.Recordset
'On Error GoTo CheckErr
'    If vbNo = MsgBox("Do you want to Delete this user?", vbQuestion + vbYesNo, "Delete") Then
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

Private Sub SaveUser()
    Const mstrMethod As String = "Save User"
    Dim rst As ADODB.Recordset
    Dim mstrUserID As String
    Dim mstrSalt As String
    Dim mintIdle As Integer
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    If vbNo = MsgBox("Do you want to Save this User?", vbQuestion + vbYesNo, mstrMethod) Then
        Exit Sub
    End If
    mstrUserID = Trim(txtUserID.Text)
    If mstrUserID = "" Then
        MsgBox "Please key in User ID", vbExclamation, mstrMethod
        Exit Sub
    End If
    If Trim(txtIdle.Text) <> "" Then
        mintIdle = ConvInt(txtIdle.Text)
        If mintIdle > 3600 Or mintIdle < 0 Then
            MsgBox "Please key in Idle time between 0 to 3600", vbExclamation, mstrMethod
            Exit Sub
        End If
    Else
        mintIdle = 0
    End If
    ' Check other textbox
    If cboUserGroup.ListIndex < 0 Then Exit Sub
    'Regenerate Salt for every update
    mstrSalt = GenSalt(4)
    SQL_SELECT_ALL "UserData"
    SQL_WHERE_Text "UserID", mstrUserID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        SQL_UPDATE "UserData"
        SQL_SET_Long "UserGroup", cboUserGroup.ItemData(cboUserGroup.ListIndex)
        SQL_SET_Text "UserID", Trim(txtUserID.Text)
        SQL_SET_Text "UserName", Trim(txtName.Text)
        SQL_SET_Integer "Idle", Trim(txtIdle.Text)
'        SQL_SET_Long "UserThemeID", cboTheme.ItemData(cboTheme.ListIndex)
        If chkActive.Value = vbChecked Then
            SQL_SET_Boolean "Active", True
        Else
            SQL_SET_Boolean "Active", False
        End If
        If chkUpdatePassword.Value = vbChecked Then
            SQL_SET_Text "UserPassword", Encrypt(txtPasswordNew.Text, mstrSalt)
            SQL_SET_Text "Salt", mstrSalt
        End If
        If chkReset.Value = vbChecked Then
            SQL_SET_Long "Loginattempts", 0
        End If
        If chkChange.Value = vbChecked Then
            SQL_SET_Boolean "ChangePassword", True, False
        Else
            SQL_SET_Boolean "ChangePassword", False, False
        End If
        SQL_WHERE_Text "UserID", mstrUserID
    Else
        SQL_INSERT "UserData"
        SQLText "UserGroup"
        SQLText "UserID"
        SQLText "UserName"
        SQLText "UserPassword"
        SQLText "Salt"
        SQLText "Idle"
        SQLText "Active"
        SQLText "Loginattempts"
        SQLText "ChangePassword", False
        SQL_VALUES
        SQLData_Long cboUserGroup.ItemData(cboUserGroup.ListIndex)
        SQLData_Text CheckInput(txtUserID.Text)
        SQLData_Text CheckInput(txtName.Text)
        SQLData_Text Encrypt(txtPasswordNew.Text, mstrSalt)
        SQLData_Text mstrSalt
        SQLData_Integer mintIdle
        If chkActive.Value = vbChecked Then
            SQLData_Boolean True
        Else
            SQLData_Boolean False
        End If
        SQLText "0"
        If chkChange.Value = vbChecked Then
            SQLData_Boolean True, False
        Else
            SQLData_Boolean False, False
        End If
        SQL_Close_Bracket
    End If
    CloseRS rst
    'Set rst = QuerySQL(gstrSQL, lngRecordsAffected)
    'CloseRS rst
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    'ResetFields
    ListUsers
    gintUserIdle = mintIdle
    MsgBox "User Account saved!", vbInformation, mstrMethod
    ' Recommend to log out and log in is required
    ' Set User Theme ID
'    gintUserThemeID = UserThemeID(gstrUserID)
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub ResetFields()
    Const mstrMethod As String = "ResetFields"
On Error GoTo CheckErr
    cboUserGroup.ListIndex = cboUserGroup.ListCount - 1
    txtUserID.Text = ""
    txtName.Text = ""
    txtIdle.Text = ""
    txtPasswordOld.Text = ""
    txtPasswordNew.Text = ""
    txtPasswordConfirm.Text = ""
    chkUpdatePassword.Value = 0
    chkActive.Value = 0
    chkReset.Value = 0
    chkReset.Caption = "Reset login attempts (0)"
    chkChange.Value = 0
'    tbrMenu.Buttons("DELETE").Enabled = False
'    txtName.SelStart = Len(txtName.Text)
'    txtName.SetFocus
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub
