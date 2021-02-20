VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFindCustomer 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Customer"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   2280
   ClientWidth     =   15375
   Icon            =   "frmFindCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3705
      ScaleWidth      =   14865
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5640
      Width           =   14895
      Begin MSComctlLib.ListView lvBookings 
         Height          =   3255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5741
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Booking No"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Room Type"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   1940
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   7560
      ScaleHeight     =   3345
      ScaleWidth      =   7545
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7575
      Begin MSComctlLib.ListView lvCustomers 
         Height          =   2895
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5106
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Passport / IC No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Guest Name"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country / Origin"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Contact No"
            Object.Width           =   2469
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   240
      ScaleHeight     =   3345
      ScaleWidth      =   7065
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7095
      Begin VB.TextBox txtGuestName 
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
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtGuestPassport 
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
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtGuestOrigin 
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
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtGuestContact 
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
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1680
         Width           =   3855
      End
      Begin VB.ComboBox cboSQL2 
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
         ItemData        =   "frmFindCustomer.frx":0EB2
         Left            =   240
         List            =   "frmFindCustomer.frx":0EBC
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cboSQL1 
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
         ItemData        =   "frmFindCustomer.frx":0EC9
         Left            =   240
         List            =   "frmFindCustomer.frx":0ED3
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboSQL3 
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
         ItemData        =   "frmFindCustomer.frx":0EE0
         Left            =   240
         List            =   "frmFindCustomer.frx":0EEA
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpBookingDateFrom 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2280
         Width           =   3135
         _ExtentX        =   5530
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
         CalendarTitleBackColor=   -2147483635
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   59441155
         CurrentDate     =   41915
      End
      Begin MSComCtl2.DTPicker dtpBookingDateTo 
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   2760
         Width           =   3135
         _ExtentX        =   5530
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
         CalendarTitleBackColor=   -2147483635
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   59441155
         CurrentDate     =   41915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Passport / IC No"
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
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country / Origin"
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
         Left            =   1200
         TabIndex        =   13
         ToolTipText     =   "Auto log-out when idle (secs)"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
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
         Left            =   1200
         TabIndex        =   14
         ToolTipText     =   "Auto log-out when idle (secs)"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Date From"
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
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Date To"
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
         TabIndex        =   16
         Top             =   2760
         Width           =   2655
      End
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Height          =   810
      Left            =   0
      TabIndex        =   17
      Top             =   1200
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   1429
      ButtonWidth     =   1984
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
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
            Caption         =   "Find (Ctrl+F)"
            Key             =   "FIND"
            Object.ToolTipText     =   "Find Customers (Ctrl+F)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print (Ctrl+P)"
            Key             =   "PRINT"
            Object.ToolTipText     =   "Print Transaction History (Ctrl+P)"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmFindCustomer.frx":0EF7
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   18
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFindCustomer.frx":1211
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFindCustomer.frx":1AEB
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFindCustomer.frx":23DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFindCustomer.frx":270F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFindCustomer.frx":33E9
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFindCustomer.frx":3CC3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmFindCustomer.frx":4B85
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
      TabIndex        =   22
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
      TabIndex        =   21
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
Attribute VB_Name = "frmFindCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Created On : 08/01/2015
' Descriptions : 1)
'
Option Explicit
Private Const mstrModule As String = "Find Customer"
Dim intTick As Integer
Dim strGuestPassport As String
Dim strGuestContact As String

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
    ResetFields
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
    ElseIf KeyCode = vbKeyF And (Shift And vbCtrlMask) And tbrMenu.Buttons("FIND").Enabled Then
        ListCustomers
    ElseIf KeyCode = vbKeyP And (Shift And vbCtrlMask) And tbrMenu.Buttons("PRINT").Enabled Then
        If lvCustomers.ListItems.Count > 0 Then
            strGuestPassport = lvCustomers.SelectedItem.Text
            strGuestContact = Trim(txtGuestContact.Text)
            PrintHistory strGuestPassport, strGuestContact
        End If
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

Private Sub lvCustomers_Click()
    Const mstrMethod As String = "lvCustomers_Click"
On Error GoTo CheckErr
    If lvCustomers.ListItems.Count = 0 Then
        Exit Sub
    End If
    ListBooking lvCustomers.SelectedItem.Text, lvCustomers.SelectedItem.ListSubItems(1).Text, lvCustomers.SelectedItem.ListSubItems(2).Text, lvCustomers.SelectedItem.ListSubItems(3).Text
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrModule 'mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub lvCustomers_KeyUp(KeyCode As Integer, Shift As Integer)
    Const mstrMethod As String = "lvCustomers_KeyUp"
On Error GoTo CheckErr
    lvCustomers_Click
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
        Case "FIND"
            ListCustomers
        Case "PRINT"
            If lvCustomers.ListItems.Count > 0 Then
                strGuestPassport = lvCustomers.SelectedItem.Text
                strGuestContact = Trim(txtGuestContact.Text)
                PrintHistory strGuestPassport, strGuestContact
            End If
        Case Else 'Default
            Exit Sub
    End Select
End Sub

'Private Sub PopulateGroupName()
'    Const mstrMethod As String = "PopulateGroupName"
'    Dim rst As ADODB.Recordset
'    Dim i As Integer
'On Error GoTo CheckErr
'    OpenDB
'    Set rst = GetList("UserGroup", "GroupID", "GroupName", "Active = TRUE", "SecurityLevel", False)
'    cboUserGroup.Clear
'    While Not rst.EOF
'        cboUserGroup.AddItem rst("GroupName").Value
'        cboUserGroup.ItemData(i) = rst("GroupID").Value
'        i = i + 1
'        rst.MoveNext
'    Wend
'    CloseRS rst
'    CloseDB
'    Exit Sub
'CheckErr:
'    CloseRS rst
'    'CloseDB
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
'    'LogErrorText "Error", mstrMethod, Err.Description
'    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
'End Sub

Private Sub ListCustomers()
    Const mstrMethod As String = "ListCustomers"
    Dim rst As ADODB.Recordset
    Dim blnAND As Boolean
    Dim List As ListItem
    'Dim LSI As ListSubItem
    'Dim i As Integer
On Error GoTo CheckErr
    lvCustomers.ListItems.Clear
    SQL_SELECT
    SQLText "DISTINCT", False
    SQLText "GuestPassport"
    SQLText "GuestName"
    SQLText "GuestOrigin"
    SQLText "GuestContact", False
    SQL_FROM "Booking"
    If Trim(txtGuestName.Text) <> "" Then
        'SQL_WHERE_Text "GuestName", Trim(txtGuestName.Text)
        SQL_WHERE_LIKE_Text "GuestName", Trim(txtGuestName.Text)
        blnAND = True
    End If
    If Trim(txtGuestPassport.Text) <> "" Then
        If blnAND = True Then
            If cboSQL1.Text = "AND" Then
                SQLText "AND GuestPassport LIKE '%" & Trim(txtGuestPassport.Text) & "%'", False
            Else
                SQLText "OR GuestPassport LIKE '%" & Trim(txtGuestPassport.Text) & "%'", False
            End If
        Else
            SQLText "WHERE GuestPassport LIKE '%" & Trim(txtGuestPassport.Text) & "%'", False
            blnAND = True
        End If
    End If
    If Trim(txtGuestOrigin.Text) <> "" Then
        If blnAND = True Then
            If cboSQL2.Text = "AND" Then
                SQLText "AND GuestOrigin LIKE '%" & Trim(txtGuestOrigin.Text) & "%'", False
            Else
                SQLText "OR GuestOrigin LIKE '%" & Trim(txtGuestOrigin.Text) & "%'", False
            End If
        Else
            SQLText "WHERE GuestOrigin LIKE '%" & Trim(txtGuestOrigin.Text) & "%'", False
            blnAND = True
        End If
    End If
    If Trim(txtGuestContact.Text) <> "" Then
        If blnAND = True Then
            If cboSQL3.Text = "AND" Then
                SQLText "AND GuestContact LIKE '%" & Trim(txtGuestContact.Text) & "%'", False
            Else
                SQLText "OR GuestContact LIKE '%" & Trim(txtGuestContact.Text) & "%'", False
            End If
        Else
            SQLText "WHERE GuestContact LIKE '%" & Trim(txtGuestContact.Text) & "%'", False
            blnAND = True
        End If
    End If
    If blnAND = True Then
        SQLText "AND (BookingDate BETWEEN #" & FormatDate(dtpBookingDateFrom.Value) & "#", False
        SQLText "AND #" & FormatDate(dtpBookingDateTo.Value) & "#)", False
        SQLText "AND Active = TRUE AND Temp = FALSE", False
    Else
        SQLText "WHERE (BookingDate BETWEEN #" & FormatDate(dtpBookingDateFrom.Value) & "#", False
        SQLText "AND #" & FormatDate(dtpBookingDateTo.Value) & "#)", False
        SQLText "AND Active = TRUE AND Temp = FALSE", False
    End If
    OpenDB
    Set rst = OpenRS(gstrSQL)
    While Not rst.EOF
        Set List = lvCustomers.ListItems.Add(, , Trim(rst!GuestPassport & " "), 0, 0)
        'List.SubItems(1) = rst!GuestPassport
        If rst!GuestName <> "" Then
            List.SubItems(1) = rst!GuestName
        Else
            List.SubItems(1) = ""
        End If
        If rst!GuestOrigin <> "" Then
            List.SubItems(2) = rst!GuestOrigin
        Else
            List.SubItems(2) = ""
        End If
        If rst!GuestContact <> "" Then
            List.SubItems(3) = rst!GuestContact
        Else
            List.SubItems(3) = ""
        End If
        lvCustomers.Refresh
        rst.MoveNext
    Wend
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

Private Sub ListBooking(strGuestPassport As String, strGuestName As String, strGuestOrigin As String, strGuestContact As String)
    Const mstrMethod As String = "ListBooking"
    Dim rst As ADODB.Recordset
    Dim List As ListItem
On Error GoTo CheckErr
    'gstrSQL = "SELECT GroupName, ID, UserID, UserName, Active, LoginAttempts, ChangePassword" & _
            " FROM UserData LEFT JOIN UserGroup ON UserData.UserGroup = UserGroup.GroupID" & _
            " WHERE UserData.ID = " & lngUserID
    lvBookings.ListItems.Clear
    SQL_SELECT
    SQLText "ID"
    SQLText "Format(ID, '100000') AS BookingNo"
    SQLText "CreatedDate"
    SQLText "RoomType"
    SQLText "Payment", False
    SQL_FROM "Booking"
    SQL_WHERE_Text "GuestPassport", strGuestPassport
    'SQLText "AND GuestName = '" & strGuestName & "'", False
    'SQL_OR_LIKE_Text "GuestOrigin", strGuestOrigin
    'SQL_OR_LIKE_Text "GuestContact", strGuestContact
    If strGuestContact <> "" Then
        If cboSQL3.Text = "AND" Then
            SQL_AND_LIKE_Text "GuestContact", strGuestContact
        Else
            SQL_OR_LIKE_Text "GuestContact", strGuestContact
        End If
    End If
    SQLText "AND (BookingDate BETWEEN #" & FormatDate(dtpBookingDateFrom.Value) & "#", False
    SQLText "AND #" & FormatDate(dtpBookingDateTo.Value) & "#)", False
    SQLText "AND Active = TRUE AND Temp = FALSE", False
    OpenDB
    Set rst = OpenRS(gstrSQL)
    While Not rst.EOF
        Set List = lvBookings.ListItems.Add(, , rst!ID, 0, 0)
        If rst!BookingNo <> "" Then
            List.SubItems(1) = rst!BookingNo
        Else
            List.SubItems(1) = ""
        End If
        If rst!CreatedDate <> "" Then
            List.SubItems(2) = rst!CreatedDate
        Else
            List.SubItems(2) = ""
        End If
        If rst!RoomType <> "" Then
            List.SubItems(3) = rst!RoomType
        Else
            List.SubItems(3) = ""
        End If
        If rst!Payment <> "" Then
            List.SubItems(4) = FormatCurrency(rst!Payment)
        Else
            List.SubItems(4) = "0.00"
        End If
        lvBookings.Refresh
        rst.MoveNext
    Wend
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

Private Sub PrintHistory(strGuestPassport As String, Optional strGuestContact As String = "")
    Const mstrMethod As String = "Print Customer History"
    Dim rep As New frmPrint
On Error GoTo CheckErr
    gstrReportFileName = "Customer Transaction History.rpt"
    gstrReportTitle = "Customer Transaction History"
    gstrSQL = "SELECT C.CompanyName, C.StreetAddress, C.ContactNo,"
    gstrSQL = gstrSQL & " B.BookingDate, Format(B.ID, '100000') AS BookingID,"
    gstrSQL = gstrSQL & " B.GuestName, B.GuestPassport, B.GuestOrigin, B.GuestContact,"
    gstrSQL = gstrSQL & " B.GuestEmergencyContactName, B.GuestEmergencyContactNo, B.StayDuration,"
    gstrSQL = gstrSQL & " B.GuestCheckIN, B.GuestCheckOUT,"
    gstrSQL = gstrSQL & " B.RoomType, (B.Payment-B.Refund) AS [Total]"
    gstrSQL = gstrSQL & " FROM Company C, Booking B"
    gstrSQL = gstrSQL & " WHERE B.Active = TRUE AND B.GuestPassport = '" & strGuestPassport & "'"
    
    If strGuestContact <> "" Then
        If cboSQL3.Text = "AND" Then
            SQL_AND_LIKE_Text "GuestContact", strGuestContact
        Else
            SQL_OR_LIKE_Text "GuestContact", strGuestContact
        End If
    End If
    
    gstrSQL = gstrSQL & " AND B.BookingDate BETWEEN #" & FormatDate(dtpBookingDateFrom.Value) & "# AND #" & FormatDate(dtpBookingDateTo.Value) & "#"
    If QueryHasData(gstrSQL) = False Then
        gstrSQL = "SELECT CompanyName, StreetAddress, ContactNo,"
        gstrSQL = gstrSQL & " '' AS BookingDate, '' AS BookingID,"
        gstrSQL = gstrSQL & " '' AS GuestName,  '' AS GuestPassport, '' AS GuestOrigin, '' AS GuestContact,"
        gstrSQL = gstrSQL & " '' AS GuestEmergencyContactName, '' AS GuestEmergencyContactNo, 0 AS StayDuration,"
        gstrSQL = gstrSQL & " '' AS GuestCheckIN, '' AS GuestCheckOUT,"
        gstrSQL = gstrSQL & " '' AS RoomType, 0 AS Total"
        gstrSQL = gstrSQL & " FROM Company"
    End If
    rep.Caption = "CUSTOMER TRANSACTION HISTORY" ' DATE from To
    rep.Show
    Exit Sub
CheckErr:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub ResetFields()
    Const mstrMethod As String = "ResetFields"
On Error GoTo CheckErr
    txtGuestName.Text = ""
    txtGuestPassport.Text = ""
    txtGuestOrigin.Text = ""
    txtGuestContact.Text = ""
    cboSQL1.ListIndex = 0
    cboSQL2.ListIndex = 0
    cboSQL3.ListIndex = 0
    dtpBookingDateFrom.Value = YearDay1(Now)
    dtpBookingDateTo.Value = MonthDay30(Now)
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub txtGuestContact_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ListCustomers
    End If
End Sub

Private Sub txtGuestName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ListCustomers
    End If
End Sub

Private Sub txtGuestOrigin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ListCustomers
    End If
End Sub

Private Sub txtGuestPassport_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ListCustomers
    End If
End Sub
