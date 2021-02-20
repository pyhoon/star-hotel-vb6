VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportMaintain 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Maintenance"
   ClientHeight    =   9645
   ClientLeft      =   4455
   ClientTop       =   2955
   ClientWidth     =   15375
   Icon            =   "frmReportMaintain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
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
      Left            =   6120
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Details"
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
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   6000
      ScaleHeight     =   7185
      ScaleWidth      =   9105
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2160
      Width           =   9135
      Begin VB.PictureBox fraAdvanced 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4335
         ScaleWidth      =   8895
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2640
         Width           =   8895
         Begin VB.ComboBox cboDateField 
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
            ItemData        =   "frmReportMaintain.frx":08CA
            Left            =   1920
            List            =   "frmReportMaintain.frx":08DA
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   0
            Width           =   6855
         End
         Begin VB.ComboBox cboDateType 
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
            ItemData        =   "frmReportMaintain.frx":091F
            Left            =   1920
            List            =   "frmReportMaintain.frx":0938
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   6855
         End
         Begin VB.TextBox txtReportQuery 
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
            Height          =   1320
            Left            =   1920
            MaxLength       =   1000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   960
            Width           =   6855
         End
         Begin VB.TextBox txtSubQuery 
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
            Left            =   1920
            MaxLength       =   255
            TabIndex        =   10
            Top             =   2400
            Width           =   6855
         End
         Begin VB.TextBox txtReportFile 
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
            Left            =   1920
            MaxLength       =   255
            TabIndex        =   12
            Top             =   3840
            Width           =   6855
         End
         Begin VB.TextBox txtNullQuery 
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
            Height          =   855
            Left            =   1920
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   2880
            Width           =   6855
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   " Date Field"
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
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   " Report Query"
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
            Left            =   0
            TabIndex        =   34
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   " Report File"
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
            Left            =   0
            TabIndex        =   33
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   " Date Type"
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
            Left            =   0
            TabIndex        =   32
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   " Sub Query"
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
            Left            =   0
            TabIndex        =   31
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   " Null Query"
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
            Left            =   0
            TabIndex        =   30
            Top             =   2880
            Width           =   1815
         End
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
         Left            =   2040
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Value           =   1  'Checked
         Width           =   6855
      End
      Begin VB.CheckBox chkDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00303030&
         Caption         =   "Show Date"
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtReportName 
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
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   2
         Top             =   720
         Width           =   6855
      End
      Begin VB.TextBox txtReportTitle 
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
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1200
         Width           =   6855
      End
      Begin VB.TextBox txtReportAsOn1 
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
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1680
         Width           =   6855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " Report Title"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblReportID 
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
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Report Name"
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
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   " Report ID"
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
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
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
         Height          =   360
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1815
      End
   End
   Begin VB.Frame fraReport 
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
      TabIndex        =   14
      Top             =   2040
      Width           =   855
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Reports"
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
         TabIndex        =   22
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Height          =   810
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   1429
      ButtonWidth     =   2223
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close (Esc) "
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Close (Esc)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save (Ctrl+S) "
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save Report (Ctrl+S)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Expert (Ctrl+E) "
            Key             =   "EXPERT"
            Object.ToolTipText     =   "Expert Mode (Ctrl+E)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Print (Ctrl+P) "
            Key             =   "PRINT"
            Object.ToolTipText     =   "Print (Ctrl+P) "
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmReportMaintain.frx":0979
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   16
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   390
            Width           =   1620
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5160
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
               Picture         =   "frmReportMaintain.frx":0C93
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReportMaintain.frx":156D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReportMaintain.frx":2247
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmReportMaintain.frx":2B21
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7185
      ScaleWidth      =   5745
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2160
      Width           =   5775
      Begin MSComctlLib.ListView lvReports 
         Height          =   6735
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   11880
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
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmReportMaintain.frx":39E3
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
      TabIndex        =   20
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
      TabIndex        =   19
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
Attribute VB_Name = "frmReportMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
Option Explicit
Private Const mstrModule As String = "Report Maintenance"
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
    Const mstrMethod As String = "Form_Load"
On Error GoTo CheckErr
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    If UserAccessModule(MOD_REPORT_EDIT_EXPERT) = True Then
        tbrMenu.Buttons("EXPERT").Enabled = True
    Else
        tbrMenu.Buttons("EXPERT").Enabled = False
    End If
    LoadReports
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrModule 'mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const mstrMethod As String = "Form_Unload"
On Error GoTo CheckErr
    'gblnMaintainReport = False
    frmReport.Show
    frmReport.LoadReports
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrModule 'mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyS And (Shift And vbCtrlMask) Then
        SaveReport
    ElseIf KeyCode = vbKeyE And (Shift And vbCtrlMask) Then
        If tbrMenu.Buttons("EXPERT").Enabled = True Then
            fraAdvanced.Visible = Not fraAdvanced.Visible
        End If
    End If
End Sub

Private Sub lvReports_Click()
    Const mstrMethod As String = "lvReports_Click"
On Error GoTo CheckErr
    If lvReports.ListItems.Count = 0 Then
        Exit Sub
    End If
    PopulateValues lvReports.SelectedItem.Text
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrModule 'mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub lvReports_KeyUp(KeyCode As Integer, Shift As Integer)
    Const mstrMethod As String = "lvReports_KeyUp"
On Error GoTo CheckErr
    lvReports_Click
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
    Case "SAVE"
        SaveReport
    Case "EXPERT"
        fraAdvanced.Visible = Not fraAdvanced.Visible
    End Select
End Sub

Public Sub SelectReport(pintReportID As Integer)
    Const mstrMethod As String = "SelectReport"
    Dim i As Integer
    Dim blnExist As Boolean
On Error GoTo CheckErr
    With lvReports
        For i = 1 To .ListItems.Count
            If .ListItems.Item(i).Text = pintReportID Then
                '.ListIndex = i
                .ListItems(i).Selected = True
                .SelectedItem.EnsureVisible
                blnExist = True
                lvReports_Click
                Exit For
            End If
        Next
        If blnExist = False Then
            lblReportID.Caption = pintReportID
        End If
    End With
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub LoadReports()
    Const mstrMethod As String = "Load Reports"
    Dim rst As ADODB.Recordset
    Dim List As ListItem
On Error GoTo CheckErr
    SQL_SELECT_ALL "Report"
    SQL_ORDER_BY "ReportID"
    OpenDB
    Set rst = OpenRS(gstrSQL)
    lvReports.ListItems.Clear
    While Not rst.EOF
        Set List = lvReports.ListItems.Add(, "i" & rst!ReportID, rst!ReportID, 0, 0)
        List.SubItems(1) = rst!ReportID
        List.SubItems(2) = rst!ReportName1
        If rst!Active = True Then
            List.ForeColor = COL_GRAY
            List.ListSubItems(1).ForeColor = COL_GRAY
            List.ListSubItems(2).ForeColor = COL_GRAY
        Else
            List.ForeColor = COL_PINK
            List.ListSubItems(1).ForeColor = COL_PINK
            List.ListSubItems(2).ForeColor = COL_PINK
        End If
        rst.MoveNext
    Wend
    CloseRS rst
    CloseDB
    lvReports_Click
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
    CloseRS rst
    CloseDB
End Sub

Private Sub PopulateValues(plngReportID As Long)
    Const mstrMethod As String = "PopulateValues"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    SQL_SELECT_ALL "Report"
    SQL_WHERE_Long "ReportID", plngReportID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        lblReportID.Caption = rst!ReportID
        SetRecord txtReportName, rst!ReportName1
        'SetRecord txtReportName2, rst!ReportName2
        SetRecord txtReportTitle, rst!ReportTitle1
        'SetRecord txtReportTitle2, rst!ReportTitle2
            If rst!DateField1 <> "" Then
                cboDateField.Text = rst!DateField1
            Else
                cboDateField.ListIndex = 0
            End If
            If rst!DateType1 <> "" Then
                cboDateType.Text = rst!DateType1
            Else
                cboDateType.ListIndex = 0
            End If
        'SetRecord txtDateType2, rst!DateType2
        SetRecord txtReportAsOn1, rst!ReportAsOn1, False
        'SetRecord txtReportAsOn, rst!ReportAsOn2
        SetRecord txtReportFile, rst!ReportFile
        'SetRecord txtTempTable, rst!TempTable
        SetRecord txtReportQuery, rst!ReportQuery
        SetRecord txtSubQuery, rst!SubQuery
        SetRecord txtNullQuery, rst!NullQuery
        SetCheck chkDate, rst!ShowReportAsOn
        SetCheck chkActive, rst!Active
    End If
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

Private Sub SaveReport()
    Const mstrMethod As String = "Save Report"
    Dim rst As ADODB.Recordset
    Dim mlngReportID As Long
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    If vbNo = MsgBox("Do you want to Save this Report?", vbQuestion + vbYesNo, mstrMethod) Then
        Exit Sub
    End If
    If lblReportID.Caption <> "" Then
        mlngReportID = ConvLng(lblReportID.Caption)
    End If
    
    SQL_SELECT_ALL "Report"
    SQL_WHERE_Long "ReportID", mlngReportID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        SQL_UPDATE "Report"
        SQL_SET_Text "ReportName1", CheckString(txtReportName.Text)
        SQL_SET_Text "ReportTitle1", CheckString(txtReportTitle.Text)
        SQL_SET_Text "ReportAsOn1", CheckString(txtReportAsOn1.Text)
        If chkDate.Value = vbChecked Then
            SQL_SET_Boolean "ShowReportAsOn", True
        Else
            SQL_SET_Boolean "ShowReportAsOn", False
        End If
        SQL_SET_Text "DateField1", CheckString(cboDateField.Text)
        SQL_SET_Text "DateType1", CheckString(cboDateType.Text)
        SQL_SET_Text "ReportFile", CheckString(txtReportFile.Text)
        SQL_SET_Text "ReportQuery", CheckString(txtReportQuery.Text)
        SQL_SET_Text "SubQuery", CheckString(txtSubQuery.Text)
        SQL_SET_Text "NullQuery", CheckString(txtNullQuery.Text)
        If chkActive.Value = vbChecked Then
            SQL_SET_Boolean "Active", True, False
        Else
            SQL_SET_Boolean "Active", False, False
        End If
        SQL_WHERE_Long "ReportID", mlngReportID
    End If
    CloseRS rst
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    'If gblnReportSelect = True Then
    '    frmReportSelect.LoadReports
    'End If
    'MsgBox "Report is updated! (" & lngRecordsAffected & ")", vbInformation, mstrMethod
    MsgBox "Report is updated!", vbInformation, mstrMethod
    LoadReports
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
    CloseRS rst
    CloseDB
End Sub
