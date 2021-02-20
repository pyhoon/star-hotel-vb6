VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoomMaintain 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Maintenance"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15375
   Icon            =   "frmRoomMaintain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRecord 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   4440
      TabIndex        =   25
      Top             =   6960
      Width           =   1335
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Record Details"
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
         Left            =   60
         TabIndex        =   47
         Top             =   0
         Width           =   1305
      End
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Height          =   810
      Left            =   0
      TabIndex        =   24
      Top             =   1200
      Width           =   15435
      _ExtentX        =   27226
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
            Object.ToolTipText     =   "Clear Fields (Ctrl+C)"
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
            Object.ToolTipText     =   "Save Room (Ctrl+S)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Type (Ctrl+T)"
            Key             =   "EDIT"
            Object.ToolTipText     =   "Edit Type (Ctrl+T)"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmRoomMaintain.frx":08CA
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   27
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   390
            Width           =   1620
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8040
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   -2147483633
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomMaintain.frx":0BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomMaintain.frx":14BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomMaintain.frx":1DB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomMaintain.frx":20E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRoomMaintain.frx":2DBC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraRoom 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   2160
      Width           =   735
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rooms"
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
         Left            =   60
         TabIndex        =   33
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   2160
      Width           =   1300
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Details"
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
         Left            =   60
         TabIndex        =   35
         Top             =   0
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7065
      ScaleWidth      =   3945
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3975
      Begin MSComctlLib.ListView lvRooms 
         Height          =   6615
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   11668
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
            Text            =   "Room No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Location"
            Object.Width           =   3616
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   4320
      ScaleHeight     =   4545
      ScaleWidth      =   10785
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2280
      Width           =   10815
      Begin VB.TextBox txtRoomShortName 
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
         Height          =   375
         Left            =   3480
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtRoomLongName 
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
         Height          =   375
         Left            =   3480
         MaxLength       =   255
         TabIndex        =   9
         Top             =   1680
         Width           =   6855
      End
      Begin VB.TextBox txtRoomPrice 
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
         Height          =   375
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   11
         Top             =   2160
         Width           =   3255
      End
      Begin VB.ComboBox cboRoomType 
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
         ItemData        =   "frmRoomMaintain.frx":3696
         Left            =   3480
         List            =   "frmRoomMaintain.frx":3698
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   6855
      End
      Begin VB.ComboBox cboLocation 
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
         ItemData        =   "frmRoomMaintain.frx":369A
         Left            =   3480
         List            =   "frmRoomMaintain.frx":36AA
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   6855
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
         TabIndex        =   21
         Top             =   4080
         Value           =   1  'Checked
         Width           =   6855
      End
      Begin VB.TextBox txtBreakfastPrice 
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
         Height          =   375
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2640
         Width           =   3255
      End
      Begin VB.CheckBox chkBreakfast 
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
         Left            =   9480
         TabIndex        =   15
         Top             =   2640
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkMaintenance 
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
         TabIndex        =   17
         Top             =   3120
         Width           =   6855
      End
      Begin VB.CheckBox chkHousekeeping 
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
         TabIndex        =   19
         Top             =   3600
         Width           =   6855
      End
      Begin VB.Label Label01 
         BackStyle       =   0  'Transparent
         Caption         =   "Room No *"
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
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label04 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Description"
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
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label02 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type *"
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
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label03 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Location *"
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
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label05 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Price (MYR)"
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
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label lblRoomID 
         BackStyle       =   0  'Transparent
         Caption         =   "Room ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Active"
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
         TabIndex        =   20
         Top             =   4080
         Width           =   3015
      End
      Begin VB.Label Label08 
         BackStyle       =   0  'Transparent
         Caption         =   "Under Maintenance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF7929&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label07 
         BackStyle       =   0  'Transparent
         Caption         =   "Breakfast Included"
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
         Height          =   375
         Left            =   7320
         TabIndex        =   14
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label06 
         BackStyle       =   0  'Transparent
         Caption         =   "Breakfast Price (MYR)"
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
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label Label09 
         BackStyle       =   0  'Transparent
         Caption         =   "Under Housekeeping"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F900D5&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label lblBookingID 
         BackStyle       =   0  'Transparent
         Caption         =   "Booking ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   9000
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   4320
      ScaleHeight     =   2265
      ScaleWidth      =   10785
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   7080
      Width           =   10815
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Created Date"
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
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Created By"
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
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified Date"
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
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified By"
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
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lblLastModifiedBy 
         BackStyle       =   0  'Transparent
         Caption         =   "<Unknown>"
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
         Height          =   375
         Left            =   3480
         TabIndex        =   42
         Top             =   1320
         Width           =   6855
      End
      Begin VB.Label lblLastModifiedDate 
         BackStyle       =   0  'Transparent
         Caption         =   "<Unknown>"
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
         Height          =   375
         Left            =   3480
         TabIndex        =   41
         Top             =   960
         Width           =   6855
      End
      Begin VB.Label lblCreatedBy 
         BackStyle       =   0  'Transparent
         Caption         =   "<Unknown>"
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
         Height          =   375
         Left            =   3480
         TabIndex        =   40
         Top             =   600
         Width           =   6855
      End
      Begin VB.Label lblCreatedDate 
         BackStyle       =   0  'Transparent
         Caption         =   "<Unknown>"
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
         Height          =   375
         Left            =   3480
         TabIndex        =   39
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Occupied Date"
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
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1680
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblLastOccupiedDate 
         BackStyle       =   0  'Transparent
         Caption         =   "<Unknown>"
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
         Height          =   375
         Left            =   3480
         TabIndex        =   37
         Top             =   1680
         Visible         =   0   'False
         Width           =   6855
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
      TabIndex        =   31
      Top             =   900
      Width           =   5295
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
      TabIndex        =   30
      Top             =   240
      Width           =   15135
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmRoomMaintain.frx":36D2
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
Attribute VB_Name = "frmRoomMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
Option Explicit
Private Const mstrModule As String = "Room Maintenance"
'Private Const COL_YELLOW = &HD6FF&     ' &HEAFF&     ' &HFFFF&
'Private Const COL_GREEN = &H76E600     ' &HFF00&
'Private Const COL_BLUE = &HFF7929      ' &HFF0000
'Private Const COL_RED = &H5700F5       ' &HFF&
'Private Const COL_PURPLE = &HF900D5    ' &HC000C0
'Private Const COL_GRAY = &H505050
Private Const COL_YELLOW = &HEAFF&
Private Const COL_GREEN = &H76E600
Private Const COL_BLUE = &HFF7929
Private Const COL_RED = &H4417FF
Private Const COL_PURPLE = &HF900D5
'Private Const COL_GRAY = &HE0E0E0
Private Const COL_PINK = &H8080FF
Private Const COL_BLACK = &H0&
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

' Logic :
' If Status = "Booked" or "Occupied" Then cannot change Status = "Maintenance" or "Free"
Public Sub Form_Load()
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    lblRoomID.Caption = "0"
    tbrMenu.Buttons("SAVE").Enabled = False
    PopulateRoomType
    ListRooms
    lvRooms_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With frmDashboard
        .SetButtonProperties
        .ShowSummary1
        .ShowSummary2 ' No need
        .ShowSummary3 ' No need
        .ShowSummary4
        .ShowSummary5
        .Show
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyC And (Shift And vbCtrlMask) And tbrMenu.Buttons("CLEAR").Enabled Then
        ResetFields
    ElseIf KeyCode = vbKeyR And (Shift And vbCtrlMask) And tbrMenu.Buttons("RESET").Enabled Then
        lvRooms_Click
    ElseIf KeyCode = vbKeyS And (Shift And vbCtrlMask) And tbrMenu.Buttons("SAVE").Enabled Then
        SaveRecord
    ElseIf KeyCode = vbKeyT And (Shift And vbCtrlMask) And tbrMenu.Buttons("EDIT").Enabled Then
        frmRoomTypeMaintain.SelectRoomType 1
        frmRoomTypeMaintain.Show
        Me.Hide
    End If
End Sub

Private Sub chkMaintenance_Click()
    If chkMaintenance.Value = vbChecked Then
        If chkHousekeeping.Value = vbChecked Then
            chkHousekeeping.Value = vbUnchecked
        End If
    End If
End Sub

Private Sub chkHousekeeping_Click()
    If chkHousekeeping.Value = vbChecked Then
        If chkMaintenance.Value = vbChecked Then
            chkMaintenance.Value = vbUnchecked
        End If
    End If
End Sub

Private Sub lvRooms_Click()
    If lvRooms.ListItems.Count = 0 Then
        Exit Sub
    Else
        If lvRooms.SelectedItem.Text > 0 Then
            PopulateValues lvRooms.SelectedItem.Text
        End If
    End If
End Sub

Private Sub lvRooms_KeyUp(KeyCode As Integer, Shift As Integer)
    lvRooms_Click
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "CLOSE"
            Unload Me
        Case "CLEAR"
            ResetFields
        Case "RESET"
            lvRooms_Click
        Case "SAVE"
            SaveRecord
        Case "EDIT"
            frmRoomTypeMaintain.SelectRoomType 1
            frmRoomTypeMaintain.Show
            Me.Hide
        Case Else 'Default
            Exit Sub
    End Select
End Sub

Private Sub PopulateRoomType()
    Const mstrMethod As String = "PopulateRoomType"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "TypeShortName"
    SQLText "Active", False
    SQL_FROM "RoomType"
    SQL_WHERE_Boolean "Active", True
    OpenDB
    Set rst = OpenRS(gstrSQL)
    'ResetFields
    cboRoomType.Clear
    While Not rst.EOF
        cboRoomType.AddItem Trim(rst("TypeShortName").Value)
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

Public Sub PopulateValues(intRoomID As Integer)
    Const mstrMethod As String = "PopulateValues"
    Dim rst As ADODB.Recordset
    Dim i As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "BookingID"
    SQLText "RoomShortName"
    SQLText "RoomLongName"
    SQLText "RoomStatus"
    SQLText "RoomType"
    SQLText "RoomLocation"
    'SQLText "Maintenance"
    SQLText "RoomPrice"
    SQLText "Breakfast"
    SQLText "BreakfastPrice"
    'SQLText "RoomPreviousPrice"
    SQLText "CreatedDate"
    SQLText "CreatedBy"
    SQLText "LastModifiedDate"
    SQLText "LastModifiedBy"
    SQLText "Active", False
    SQL_FROM "Room"
    SQL_WHERE_Integer "ID", intRoomID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    ''''''''''''''''''ResetFields
    If Not rst.EOF Then
        ' Room Details
        lblRoomID.Caption = intRoomID
        If rst!BookingID <> "" Then
            lblBookingID.Caption = rst!BookingID
        Else
            lblBookingID.Caption = "0"
        End If
        If rst!RoomShortName <> "" Then
            txtRoomShortName.Text = rst!RoomShortName
        Else
            txtRoomShortName.Text = ""
        End If
        If rst!RoomLongName <> "" Then
            txtRoomLongName.Text = rst!RoomLongName
        Else
            txtRoomLongName.Text = ""
        End If
        'cboStatus.Text = rst("RoomStatus").Value
        If rst!RoomType <> "" Then
            For i = 0 To cboRoomType.ListCount - 1
                If cboRoomType.List(i) = rst!RoomType Then
                    cboRoomType.ListIndex = i
                    Exit For
                End If
            Next
        Else
            cboRoomType.ListIndex = -1
        End If
        'cboRoomType.Text = rst("RoomType").Value
        If rst!RoomLocation <> "" Then
            For i = 0 To cboLocation.ListCount - 1
                If cboLocation.List(i) = rst!RoomLocation Then
                    cboLocation.ListIndex = i
                    Exit For
                End If
            Next
        Else
            cboLocation.ListIndex = -1
        End If
        'cboLocation.Text = rst("RoomLocation").Value
        txtRoomPrice.Text = FormatCurrency(rst("RoomPrice").Value)
        txtBreakfastPrice.Text = FormatCurrency(rst("BreakfastPrice").Value)
        If rst("Breakfast").Value = True Then
            chkBreakfast.Value = vbChecked
        Else
            chkBreakfast.Value = vbUnchecked
        End If
        If rst("RoomStatus").Value = "Booked" Or rst("RoomStatus").Value = "Occupied" Then
            ' Disable edit
            txtRoomShortName.Enabled = False
            txtRoomLongName.Enabled = False
            txtRoomPrice.Enabled = False
            txtBreakfastPrice.Enabled = False
            cboRoomType.Enabled = False
            cboLocation.Enabled = False
            chkBreakfast.Enabled = False
            chkMaintenance.Enabled = False
            chkHousekeeping.Enabled = False
            chkActive.Enabled = False
            tbrMenu.Buttons("SAVE").Enabled = False
        Else
            txtRoomShortName.Enabled = True
            txtRoomLongName.Enabled = True
            txtRoomPrice.Enabled = True
            txtBreakfastPrice.Enabled = True
            cboRoomType.Enabled = True
            cboLocation.Enabled = True
            chkBreakfast.Enabled = True
            chkMaintenance.Enabled = True
            chkHousekeeping.Enabled = True
            chkActive.Enabled = True
            tbrMenu.Buttons("SAVE").Enabled = True
        End If
        If rst("RoomStatus").Value = "Maintenance" Then
            chkMaintenance.Value = vbChecked
        Else
            chkMaintenance.Value = vbUnchecked
        End If
        If rst("RoomStatus").Value = "Housekeeping" Then
            chkHousekeeping.Value = vbChecked
        Else
            chkHousekeeping.Value = vbUnchecked
        End If
        If rst("Active").Value = True Then
            chkActive.Value = vbChecked
        Else
            chkActive.Value = vbUnchecked
            chkMaintenance.Enabled = False
            chkHousekeeping.Enabled = False
        End If
        ' Record Details
        If rst("CreatedDate").Value <> "" Then
            lblCreatedDate.Caption = FormatDateAndTime(rst("CreatedDate").Value)
        Else
            lblCreatedDate.Caption = ""
        End If
        If rst("CreatedBy").Value <> "" Then
            lblCreatedBy.Caption = rst("CreatedBy").Value
        Else
            lblCreatedBy.Caption = ""
        End If
        'lblCreatedBy.Caption = rst("CreatedBy").Value
        If rst("LastModifiedDate").Value <> "" Then
            lblLastModifiedDate.Caption = FormatDateAndTime(rst("LastModifiedDate").Value)
        Else
            lblLastModifiedDate.Caption = "<Never>"
        End If
        If rst("LastModifiedBy").Value <> "" Then
            lblLastModifiedBy.Caption = rst("LastModifiedBy").Value
        Else
            lblLastModifiedBy.Caption = "<Never>"
        End If
        'lblRoomPreviousPrice.Caption = rst("RoomPreviousPrice").Value
    Else
        'MsgBox "Room is not exist!", vbInformation, mstrModule 'mstrMethod
        ' Room Details
        lblRoomID.Caption = intRoomID
        txtRoomShortName.Text = ""
        txtRoomLongName.Text = ""
        'cboStatus.ListIndex = -1
        cboRoomType.ListIndex = -1
        cboLocation.ListIndex = -1
        txtRoomPrice.Text = "" '"0.00"
        txtBreakfastPrice.Text = "" '"0.00"
        chkBreakfast.Value = vbChecked
        chkMaintenance.Value = vbUnchecked
        chkHousekeeping.Value = vbUnchecked
        chkActive.Value = vbChecked
        ' Record Details
        lblCreatedDate.Caption = "<New>"
        lblCreatedBy.Caption = "<New>"
        lblLastModifiedDate.Caption = "<New>"
        lblLastModifiedBy.Caption = "<New>"
        If intRoomID = 0 Then
            tbrMenu.Buttons("SAVE").Enabled = False
        Else
            tbrMenu.Buttons("SAVE").Enabled = True
        End If
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

Private Sub ListRooms()
    Const mstrMethod As String = "ListRooms"
    Dim rst As ADODB.Recordset
    Dim List As ListItem
    Dim LSI As ListSubItem
    Dim i As Integer
On Error GoTo CheckErr
    lvRooms.ListItems.Clear
    SQL_SELECT
    SQLText "ID"
    SQLText "RoomShortName"
    'SQLText "RoomLongName"
    SQLText "RoomStatus"
    'SQLText "RoomType"
    SQLText "RoomLocation"
    'SQLText "Maintenance"
    SQLText "Active", False
    SQL_FROM "Room"
    OpenDB
    Set rst = OpenRS(gstrSQL)
    While Not rst.EOF
        Set List = lvRooms.ListItems.Add(, "i" & rst!ID, rst!ID, 0, 0)
        If rst!RoomShortName <> "" Then
            List.SubItems(1) = rst!RoomShortName
        Else
            List.SubItems(1) = ""
        End If
        'List.SubItems(2) = rst!RoomLongName
        If rst!RoomLocation <> "" Then
            List.SubItems(2) = rst!RoomLocation
        Else
            List.SubItems(2) = ""
        End If
'        If rst!Active = True Then
'            List.SubItems(5) = "Yes"
'        Else
'            List.SubItems(5) = ""
'        End If
        If rst!Active = False Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_PINK
                LSI.Bold = True
            Next
        ElseIf rst!RoomStatus = "Maintenance" Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_BLUE       ' vbBlue
                LSI.Bold = True
            Next
        ElseIf rst!RoomStatus = "Housekeeping" Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_PURPLE     ' vbPurple
                LSI.Bold = True
            Next
        ElseIf rst!RoomStatus = "Open" Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_GREEN      ' vbGreen
                LSI.Bold = True
            Next
        ElseIf rst!RoomStatus = "Booked" Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_YELLOW      ' vbYellow
            Next
        ElseIf rst!RoomStatus = "Occupied" Then
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_RED         ' vbRed
            Next
        Else
            For i = 1 To 2
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = COL_BLACK      ' vbBlack
            Next
        End If
        lvRooms.Refresh
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

Public Sub SelectRoom(pintRoomID As Integer)
    Const mstrMethod As String = "SelectRoom"
    Dim i As Integer
    Dim blnExist As Boolean
On Error GoTo CheckErr
    With lvRooms
        For i = 1 To .ListItems.Count
            If .ListItems.Item(i).Text = pintRoomID Then
                '.ListIndex = i
                .ListItems(i).Selected = True
                .SelectedItem.EnsureVisible
                blnExist = True
                lvRooms_Click
                Exit For
            End If
        Next
        If blnExist = False Then
            lblRoomID.Caption = pintRoomID ' ""
            ResetFields
        End If
    End With
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub SaveRecord()
    Const mstrMethod As String = "Save Room"
    Dim rst As ADODB.Recordset
    Dim intRoomID As Integer
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    intRoomID = CInt(lblRoomID.Caption)
    If intRoomID < 1 Then
        Exit Sub
    End If
    If vbNo = MsgBox("Do you want to Save this Room?", vbQuestion + vbYesNo, mstrMethod) Then
        Exit Sub
    End If
    If Trim(txtRoomShortName.Text) = "" Then
        MsgBox "Please key in Room No", vbExclamation, mstrMethod
        txtRoomShortName.SetFocus
        Exit Sub
    End If
    If cboRoomType.ListIndex < 0 Then
        MsgBox "Please select Room Type", vbExclamation, mstrMethod
        cboRoomType.SetFocus
        Exit Sub
    End If
    If cboLocation.ListIndex < 0 Then
        MsgBox "Please select Location", vbExclamation, mstrMethod
        cboLocation.SetFocus
        Exit Sub
    End If
    SQL_SELECT_ALL "Room"
    SQL_WHERE_Integer "ID", intRoomID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        'Copy to LogRoom
        SQL_INSERT "LogRoom"
        SQLText "BookingID", True, False
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
        'SQLText "Maintenance"
        SQLText "Active", False
        SQL_Close_Bracket
        SQLText "SELECT", False
        'SQL_VALUES
        SQLText "BookingID"
        SQLText "ID"
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
        'SQLText "Maintenance"
        SQLText "Active", False
        SQL_FROM "Room"
        'SQL_WHERE_Text "RoomShortName", mstrRoomNo
        SQL_WHERE_Integer "ID", intRoomID
        QuerySQL gstrSQL, lngRecordsAffected
                       
        'Update table Room
        SQL_UPDATE "Room"
        SQL_SET_Text "RoomShortName", Trim(txtRoomShortName.Text)
        SQL_SET_Text "RoomLongName", Trim(txtRoomLongName.Text)
        'SQL_SET_Text "RoomStatus", Trim(cboStatus.Text)
        SQL_SET_Text "RoomType", Trim(cboRoomType.Text)
        SQL_SET_Text "RoomLocation", Trim(cboLocation.Text)
        SQL_SET_Double "RoomPrice", Val(txtRoomPrice.Text)
        If chkBreakfast.Value = vbChecked Then
            SQL_SET_Boolean "Breakfast", True
        Else
            SQL_SET_Boolean "Breakfast", False
        End If
        SQL_SET_Double "BreakfastPrice", Val(txtBreakfastPrice.Text)
        ' Record Details
        SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
        SQL_SET_Text "LastModifiedBy", gstrUserID
        If chkMaintenance.Value = vbChecked Then
            SQL_SET_Text "RoomStatus", "Maintenance"
        ElseIf chkHousekeeping.Value = vbChecked Then
            SQL_SET_Text "RoomStatus", "Housekeeping"
        ElseIf chkHousekeeping.Value = vbUnchecked Then
            SQL_SET_Text "RoomStatus", "Open"
            SQL_SET_Long "BookingID", 0
        Else
            SQL_SET_Text "RoomStatus", "Open"
        End If
        If chkActive.Value = vbChecked Then
            SQL_SET_Boolean "Active", True, False
        Else
            SQL_SET_Boolean "Active", False, False
        End If
        SQL_WHERE_Integer "ID", intRoomID
        'SQLText "AND (RoomStatus <> 'Booked' OR RoomStatus <> 'Occupied')", False
    Else
        SQL_INSERT "Room"
        SQLText "ID", True, False
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
        SQLText "Active", False
        SQL_VALUES
        SQLData_Integer intRoomID, True, False
        SQLData_Text Trim(txtRoomShortName.Text)
        SQLData_Text Trim(txtRoomLongName.Text)
        If chkMaintenance.Value = vbChecked Then
            SQLData_Text "Maintenance"
        ElseIf chkHousekeeping.Value = vbChecked Then
            SQLData_Text "Housekeeping"
        Else
            SQLData_Text "Open"
        End If
        SQLData_Text Trim(cboRoomType.Text)
        SQLData_Text Trim(cboLocation.Text)
        SQLData_Double Val(txtRoomPrice.Text)
        If chkBreakfast.Value = vbChecked Then
            SQLData_Boolean True
        Else
            SQLData_Boolean False
        End If
        SQLData_Double Val(txtBreakfastPrice.Text)
        SQLData_DateTime FormatDateAndTime(Now)
        SQLData_Text gstrUserID
        If chkActive.Value = vbChecked Then
            SQLData_Boolean True, False
        Else
            SQLData_Boolean False, False
        End If
        SQL_Close_Bracket
    End If
    CloseRS rst
    QuerySQL gstrSQL, lngRecordsAffected
    
'    'Update Room BookingID = 0 if status = Housekeeping -> Free
'    If chkHousekeeping.Value = vbUnchecked Then
'        SQL_UPDATE "Room"
'        SQL_SET_Long "BookingID", 0
'        SQL_SET_Text "RoomStatus", "Open", False
'        SQL_WHERE_Integer "ID", intRoomID
'        SQLText "AND RoomStatus = 'Housekeeping'", False
'        QuerySQL gstrSQL, lngRecordsAffected
'        CloseDB
'        MsgBox "Room is free!", vbInformation, mstrMethod
'        ListRooms
'        SelectRoom intRoomID
'        lvRooms_Click
'        Exit Sub
'    'Else
'
'    End If
    CloseDB
    MsgBox "Room is saved!", vbInformation, mstrMethod
    'ResetFields
    ListRooms
    SelectRoom intRoomID
    lvRooms_Click
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
    Dim intRoomID As Integer
On Error GoTo CheckErr
    intRoomID = CInt(lblRoomID.Caption)
    cboRoomType.ListIndex = -1
    Select Case intRoomID
        Case 0 To 22
            cboLocation.Text = "Level 1"
        Case 23 To 39
            cboLocation.Text = "Level 2"
        Case 40 To 50
            cboLocation.Text = "Level 3"
        Case 51 To 61
            cboLocation.Text = "Level 4"
        Case Else
            cboLocation.ListIndex = -1
    End Select
    txtRoomShortName.Text = ""
    txtRoomLongName.Text = ""
    txtRoomPrice.Text = "" '"0.00"
    chkBreakfast.Value = vbChecked
    txtBreakfastPrice.Text = "" '"0.00"
    chkMaintenance.Value = vbUnchecked
    chkActive.Value = vbChecked
    'lblRoomID.Caption = "0"
    lblCreatedDate.Caption = "<New>"
    lblCreatedBy.Caption = "<New>"
    lblLastModifiedDate.Caption = "<New>"
    lblLastModifiedBy.Caption = "<New>"
    lvRooms.HideSelection = True
    'txtRoomShortName.SetFocus
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub
