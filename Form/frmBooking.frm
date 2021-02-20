VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBooking 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Booking"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15375
   Icon            =   "frmBooking.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraOther 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9600
      TabIndex        =   36
      Top             =   7800
      Width           =   950
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   " Remarks"
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
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   950
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
      Left            =   9600
      TabIndex        =   35
      Top             =   2760
      Width           =   1335
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   60
         TabIndex        =   68
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame fraGuest 
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
      Left            =   360
      TabIndex        =   33
      Top             =   5640
      Width           =   1335
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Guest Details"
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
         Left            =   60
         TabIndex        =   52
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2145
      ScaleWidth      =   8985
      TabIndex        =   47
      Top             =   5760
      Width           =   9015
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
         Height          =   375
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   8
         Top             =   240
         Width           =   6375
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
         Height          =   375
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   9
         Top             =   720
         Width           =   6375
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
         Height          =   375
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   10
         Top             =   1200
         Width           =   6375
      End
      Begin VB.TextBox txtContactNo 
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
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   11
         Top             =   1680
         Width           =   6375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name *"
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
         TabIndex        =   51
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Passport / IC No *"
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
         TabIndex        =   50
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label5 
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
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   1680
         Width           =   1935
      End
   End
   Begin VB.Frame fraBooking 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      Caption         =   "Booking Details"
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
      Left            =   360
      TabIndex        =   32
      Top             =   2760
      Width           =   1575
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Details"
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
         Left            =   60
         TabIndex        =   46
         Top             =   0
         Width           =   1455
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
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close (Esc)"
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Close (Esc)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset (Ctrl+R)"
            Key             =   "RESET"
            Object.ToolTipText     =   "Reset (Ctrl+R)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save (Ctrl+S)"
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save Booking (Ctrl+S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Void (Ctrl+D)"
            Key             =   "VOID"
            Object.ToolTipText     =   "Void Booking (Ctrl+D)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "IN (Ctrl+I)"
            Key             =   "Check-IN"
            Object.ToolTipText     =   "Check IN (Ctrl+I)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OUT (Ctrl+O)"
            Key             =   "Check-OUT"
            Object.ToolTipText     =   "Check OUT (Ctrl+O)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "T/R (F11)"
            Key             =   "TEMPORARY"
            Object.ToolTipText     =   "Print Temporary Receipt (F11)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "O/R (F12)"
            Key             =   "OFFICIAL"
            Object.ToolTipText     =   "Print Official Receipt (F12)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmBooking.frx":4888A
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   25
         Top             =   0
         Width           =   4935
         Begin VB.Timer tmrClock 
            Interval        =   1000
            Left            =   0
            Top             =   120
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
            TabIndex        =   27
            Top             =   390
            Width           =   1620
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
            TabIndex        =   26
            Top             =   120
            Width           =   3900
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9480
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   -2147483633
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBooking.frx":48BA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBooking.frx":4947E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBooking.frx":497B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBooking.frx":4A48A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBooking.frx":4B164
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBooking.frx":4BE3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBooking.frx":4C730
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBooking.frx":4D022
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraStatus 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   28
      Top             =   1900
      Width           =   15455
      Begin VB.Label lblBookingID 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblBookingNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Booking No :                                  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   14250
         TabIndex        =   31
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.Frame fraEmergency 
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
      Left            =   360
      TabIndex        =   34
      Top             =   8040
      Width           =   1855
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Emergency Contact"
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
         Left            =   60
         TabIndex        =   56
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      FillColor       =   &H00E0E0E0&
      ForeColor       =   &H00E0E0E0&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2625
      ScaleWidth      =   8985
      TabIndex        =   40
      Top             =   2880
      Width           =   9015
      Begin VB.ComboBox cboTotalGuest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmBooking.frx":4DEE4
         Left            =   3000
         List            =   "frmBooking.frx":4DEF4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cboStayDuration 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmBooking.frx":4DF1D
         Left            =   3000
         List            =   "frmBooking.frx":4DF1F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtpCheckOutDate 
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   2160
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
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   4210752
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   59375619
         CurrentDate     =   41916
      End
      Begin MSComCtl2.DTPicker dtpCheckInDate 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   1680
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
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   4210752
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   59375619
         CurrentDate     =   41915
      End
      Begin MSComCtl2.DTPicker dtpBookingDate 
         Height          =   375
         Left            =   3000
         TabIndex        =   0
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
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   4210752
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   59375619
         CurrentDate     =   41915
      End
      Begin MSComCtl2.DTPicker dtpCheckInTime 
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
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
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   255
         CalendarTitleForeColor=   4210752
         CustomFormat    =   "hh : mm  tt"
         Format          =   59375619
         UpDown          =   -1  'True
         CurrentDate     =   41915.5
      End
      Begin MSComCtl2.DTPicker dtpCheckOutTime 
         Height          =   375
         Left            =   6720
         TabIndex        =   7
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
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
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   65280
         CalendarTitleForeColor=   4210752
         CustomFormat    =   "hh : mm  tt"
         Format          =   59375619
         UpDown          =   -1  'True
         CurrentDate     =   41915.5
      End
      Begin MSComCtl2.DTPicker dtpBookingTime 
         Height          =   375
         Left            =   6720
         TabIndex        =   1
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
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
         CalendarForeColor=   4210752
         CalendarTitleBackColor=   65535
         CalendarTitleForeColor=   4210752
         CustomFormat    =   "hh : mm  tt"
         Format          =   59375619
         UpDown          =   -1  'True
         CurrentDate     =   41915.5
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Guest (Persons)"
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
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblTotalDays 
         BackStyle       =   0  'Transparent
         Caption         =   "Length of Stay (Nights)"
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
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Date"
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
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Date && Time Check-IN"
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
         TabIndex        =   42
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Date && Time Check-OUT"
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
         TabIndex        =   41
         Top             =   2160
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   8985
      TabIndex        =   53
      Top             =   8160
      Width           =   9015
      Begin VB.TextBox txtGuestEmergencyContactNo 
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
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   13
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox txtGuestEmergencyContactName 
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
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   12
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
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
         TabIndex        =   54
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   9480
      ScaleHeight     =   4665
      ScaleWidth      =   5625
      TabIndex        =   57
      Top             =   2880
      Width           =   5655
      Begin VB.TextBox txtDeposit 
         Alignment       =   2  'Center
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
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   19
         Text            =   "20.00"
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtPayment 
         Alignment       =   2  'Center
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
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   21
         Top             =   3600
         Width           =   2655
      End
      Begin VB.TextBox txtRefund 
         Alignment       =   2  'Center
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
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   22
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Room No"
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
         Height          =   225
         Left            =   240
         TabIndex        =   67
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblRoomID 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   1320
         TabIndex        =   66
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Due (MYR)"
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
         TabIndex        =   65
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate (MYR)"
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
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Height          =   225
         Left            =   240
         TabIndex        =   63
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Type"
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
         Height          =   225
         Left            =   240
         TabIndex        =   62
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblTotalDue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2520
         TabIndex        =   20
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label lblRoomNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblRoomType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2520
         TabIndex        =   15
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblLocation 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2520
         TabIndex        =   16
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblRoomPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   2520
         TabIndex        =   17
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit (MYR)"
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
         TabIndex        =   61
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment (MYR)"
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
         TabIndex        =   60
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblSubTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00505050&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Left            =   2520
         TabIndex        =   18
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total (MYR)"
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
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Refund (MYR)"
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
         TabIndex        =   58
         Top             =   4080
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   9480
      ScaleHeight     =   1545
      ScaleWidth      =   5625
      TabIndex        =   69
      Top             =   7920
      Width           =   5655
      Begin VB.TextBox txtRemarks 
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
         Height          =   1095
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label lblBreakfastPrice 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "10.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4440
         TabIndex        =   71
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblBreakfast 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         TabIndex        =   70
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
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
      TabIndex        =   39
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
      TabIndex        =   38
      Top             =   240
      Width           =   15135
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmBooking.frx":4DF21
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
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 07/01/2015
' Descriptions : 1) Change Clear button to Reset
'                2) Rearrange menu buttons and change labels and shortcut keys
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
'
' Modified On : 02/11/2014
' Descriptions : 1) Room is set to Maintenance when CHECK_OUT
'                2) Room is set to Maintenance when SaveBooking if status = "Closed"
'                3) Check is payment is full if Close room or CHECK_OUT
'
Option Explicit
Private Const mstrModule As String = "Booking"
'Private Const COL_YELLOW = &HFFFF&
'Private Const COL_GREEN = &HFF00&
'Private Const COL_BLUE = &HFF0000
'Private Const COL_RED = &HFF&
'Private Const COL_PURPLE = &HC000C0
'Private Const COL_GRAY = &HC0C0C0
Private Const COL_YELLOW = &HEAFF&
Private Const COL_GREEN = &H76E600
Private Const COL_BLUE = &HFF7929
Private Const COL_RED = &H4417FF
Private Const COL_PURPLE = &HF900D5
Private Const COL_GRAY = &H505050
Dim lngBookingID As Long
Dim intRoomID As Integer
Dim strStatus As String
Dim blnActive As Boolean
Dim intTick As Integer
Dim datCheckIn As Date
'Dim datCheckInStart As Date ' If datCheckIn = 2:00PM datCheckInStart = 12:00PM, If after 12:00AM datCheckInStart = 12:00PM 1 day before
Dim datCheckOut As Date
Dim intDay As Integer

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
    Dim i As Integer
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    cboTotalGuest.Clear
    For i = 1 To 6
        cboTotalGuest.AddItem i
    Next
    cboStayDuration.Clear
    For i = 1 To 10
        cboStayDuration.AddItem i
    Next
    ResetFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With frmDashboard
        .SetButtonProperties
        .ShowSummary1
        .ShowSummary2
        .ShowSummary3
        .ShowSummary4
        .Show
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyR And (Shift And vbCtrlMask) And tbrMenu.Buttons("RESET").Enabled Then
        ResetFields
        'PopulateValues lngBookingID
        SelectRoom intRoomID
    ElseIf KeyCode = vbKeyS And (Shift And vbCtrlMask) And tbrMenu.Buttons("SAVE").Enabled Then
        SumTotal
        SaveBooking
    ' VOID is not implemented since not a request
    'ElseIf KeyCode = vbKeyD And (Shift And vbCtrlMask) And tbrMenu.Buttons("VOID").Enabled Then
    '    If UserAccessModule(MOD_BOOKING_VOID) = True Then
    '        VoidBooking lngBookingID
    '    Else
    '        frmAdmin.lblBookingID.Caption = lngBookingID
    '        frmAdmin.Show vbModal
    '    End If
    ElseIf KeyCode = vbKeyI And (Shift And vbCtrlMask) And tbrMenu.Buttons("Check-IN").Enabled Then
        Check_IN
    ElseIf KeyCode = vbKeyO And (Shift And vbCtrlMask) And tbrMenu.Buttons("Check-OUT").Enabled Then
        Check_OUT
    ElseIf KeyCode = vbKeyF11 And (Shift And vbCtrlMask) And tbrMenu.Buttons("TEMPORARY").Enabled Then
        If lngBookingID <> 0 Then
            PrintReceipt "TEMPORARY"
        End If
    ElseIf KeyCode = vbKeyF12 And (Shift And vbCtrlMask) And tbrMenu.Buttons("OFFICIAL").Enabled Then
        If strStatus = "Housekeeping" Then
            PrintReceipt "OFFICIAL"
            Exit Sub
        End If
        If lngBookingID <> 0 Then
            If vbYes = MsgBox("This Room is not Checked Out." & vbCrLf & _
            "Are you sure to continue?", vbYesNo + vbQuestion, "NOT CHECK OUT") Then
            PrintReceipt "OFFICIAL"
            End If
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub cboStayDuration_Click()
    intDay = CInt(cboStayDuration.Text)
    If TimeValue(dtpCheckInTime.Value) >= TimeValue("12:00 PM") Then
        dtpCheckOutDate.Value = DateAdd("D", intDay, dtpCheckInDate.Value)
        dtpCheckOutTime.Value = "12:00 PM"
    Else
        dtpCheckOutDate.Value = DateAdd("D", intDay - 1, dtpCheckInDate.Value)
        dtpCheckOutTime.Value = "12:00 PM"
    End If
    SumTotal
End Sub

Private Sub dtpCheckInDate_Change()
    If dtpCheckInDate.Value <> DateValue(Now) Then
        dtpCheckInTime.Value = "12:00 PM"
    End If
    dtpCheckOutDate.Value = DateAdd("D", intDay, dtpCheckInDate.Value)
    dtpCheckOutTime.Value = FormatDate(dtpCheckOutDate.Value) & " 12:00 PM"
    SumTotal
End Sub

Private Sub dtpCheckOutDate_Change()
    dtpCheckInDate.Value = DateAdd("D", -intDay, dtpCheckOutDate.Value)
    SumTotal
End Sub

Private Sub dtpCheckInTime_Change()
    '
    'SumTotal
End Sub

Private Sub dtpCheckOutTime_Change()
    If TimeValue(dtpCheckInTime.Value) >= TimeValue("12:00 PM") Then
        dtpCheckOutDate.Value = DateAdd("D", intDay, dtpCheckInDate.Value)
    ElseIf TimeValue(dtpCheckInTime.Value) >= TimeValue("2:00 PM") Then
        dtpCheckOutDate.Value = DateAdd("D", intDay, dtpCheckInDate.Value)
    Else
        dtpCheckOutDate.Value = DateAdd("D", intDay - 1, dtpCheckInDate.Value)
    End If
    SumTotal
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim strTempUserID As String
    Dim strTempPassword As String
Select Case Button.Key
    Case "CLOSE"
        Unload Me
    Case "RESET"
        ResetFields
        'PopulateValues lngBookingID
        SelectRoom intRoomID
    Case "SAVE"
        SumTotal
        SaveBooking
    ' VOID is not implemented since not a request
    'Case "VOID"
    '    If UserAccessModule(MOD_BOOKING_VOID) = True Then
    '        VoidBooking lngBookingID
    '    Else
    '        frmAdmin.lblBookingID.Caption = lngBookingID
    '        frmAdmin.Show vbModal
    '    End If
    Case "Check-IN"
        Check_IN
    Case "Check-OUT"
        Check_OUT
    Case "TEMPORARY"
        If lngBookingID <> 0 Then
            PrintReceipt "TEMPORARY"
        End If
    Case "OFFICIAL"
        If strStatus = "Housekeeping" Then
            PrintReceipt "OFFICIAL"
            Exit Sub
        End If
        If lngBookingID <> 0 Then
            If vbYes = MsgBox("This Room is not Checked Out." & vbCrLf & _
            "Are you sure to continue?", vbYesNo + vbQuestion, "NOT CHECK OUT") Then
            PrintReceipt "OFFICIAL"
            End If
        End If
    Case Else
    
End Select
End Sub

Private Sub ShowStatus(pstrStatus As String)
    Select Case pstrStatus
        Case "Open"
            fraStatus.BackColor = COL_GREEN
            lblStatus.Caption = pstrStatus
            tbrMenu.Buttons("Check-IN").Enabled = False
            tbrMenu.Buttons("Check-OUT").Enabled = False
            txtRefund.Enabled = False
        Case "Booked"
            fraStatus.BackColor = COL_YELLOW
            lblStatus.Caption = pstrStatus
            tbrMenu.Buttons("Check-IN").Enabled = True
            tbrMenu.Buttons("Check-OUT").Enabled = False
            txtRefund.Enabled = False
        Case "Occupied"
            fraStatus.BackColor = COL_RED
            lblStatus.Caption = pstrStatus
            lblBookingID.ForeColor = vbWhite
            tbrMenu.Buttons("Check-IN").Enabled = False
            tbrMenu.Buttons("Check-OUT").Enabled = True
            txtRefund.Enabled = True
        Case "Housekeeping"
            fraStatus.BackColor = COL_PURPLE
            lblStatus.ForeColor = vbWhite
            lblStatus.Caption = pstrStatus
            'lblBookingID.ForeColor = vbWhite
            'lblBookingNo.Visible = False
            'lblBookingID.Visible = False
            tbrMenu.Buttons("SAVE").Enabled = False
            tbrMenu.Buttons("Check-IN").Enabled = False
            tbrMenu.Buttons("Check-OUT").Enabled = False
            ' Enable Print Buttons
            'tbrMenu.Buttons("TEMPORARY").Enabled = False
            'tbrMenu.Buttons("OFFICIAL").Enabled = False
            DisableControls
        Case "Maintenance"
            fraStatus.BackColor = COL_BLUE
            lblStatus.Caption = pstrStatus
            'lblBookingID.ForeColor = vbWhite
            lblBookingNo.Visible = False
            lblBookingID.Visible = False
            tbrMenu.Buttons("Check-IN").Enabled = False
            tbrMenu.Buttons("Check-OUT").Enabled = False
        Case "Void"
            fraStatus.BackColor = COL_GRAY
            lblStatus.Caption = pstrStatus
            tbrMenu.Buttons("Check-IN").Enabled = False
            tbrMenu.Buttons("Check-OUT").Enabled = False
            txtRefund.Enabled = False
    End Select
End Sub

Private Sub SumTotal()
    Const mstrMethod As String = "SumTotal"
    Dim intDay As Integer
    Dim dblRoomPrice As Double
    Dim dblSubTotal As Double
    Dim dblDeposit As Double
On Error GoTo CheckErr
    intDay = CInt(cboStayDuration.Text)
    If lblRoomPrice.Caption <> "" Then
        dblRoomPrice = ConvDbl(lblRoomPrice.Caption)
    Else
        dblRoomPrice = 0
    End If
    If txtDeposit.Text <> "" Then
        dblDeposit = ConvDbl(txtDeposit.Text)
    Else
        dblDeposit = 0
    End If
    dblSubTotal = intDay * dblRoomPrice
    lblSubTotal.Caption = FormatCurrency(dblSubTotal)
    lblTotalDue.Caption = FormatCurrency(dblSubTotal + dblDeposit)
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Sub PopulateValues(plngBookingID As Long)
    Const mstrMethod As String = "PopulateValues"
    Dim rst As ADODB.Recordset
    Dim intRoomID As Integer
    Dim dblSubTotal As Double
    Dim dblDeposit As Double
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "B.ID"
    SQLText "GuestName"
    SQLText "GuestPassport"
    SQLText "GuestEmergencyContactName"
    SQLText "GuestEmergencyContactNo"
    SQLText "GuestOrigin"
    SQLText "GuestContact"
    SQLText "BookingDate"
    SQLText "GuestCheckIN"
    SQLText "GuestCheckOUT"
    SQLText "TotalGuest"
    SQLText "StayDuration"
    'SQLText "Status"
    SQLText "Remarks"
    'SQLText "RoomID"
    SQLText "RoomNo"
    SQLText "R.RoomStatus"
    SQLText "B.RoomType"
    SQLText "B.RoomLocation"
    SQLText "B.RoomPrice"
    SQLText "B.Breakfast"
    SQLText "B.BreakfastPrice"
    SQLText "B.SubTotal"
    SQLText "B.Deposit"
    SQLText "B.Payment"
    SQLText "B.Refund"
    'SQLText "CreatedDate"
    'SQLText "CreatedBy"
    'SQLText "LastModifiedDate"
    'SQLText "LastModifiedBy"
    SQLText "B.Active", False
    SQL_FROM "Booking B, Room R"
    SQL_WHERE_Long "B.ID", plngBookingID
    SQLText "AND B.ID = R.BookingID", False
    OpenDB
    Set rst = OpenRS(gstrSQL)
    ''''''''''''''''''ResetFields
    If Not rst.EOF Then
        ' Booking Details
        lblBookingID.Caption = Format(rst("ID").Value, "100000")
        txtGuestName.Text = Trim(rst("GuestName").Value)
        txtGuestPassport.Text = rst("GuestPassport").Value
        txtGuestEmergencyContactName.Text = ConvText(rst("GuestEmergencyContactName").Value)
        txtGuestEmergencyContactNo.Text = ConvText(rst("GuestEmergencyContactNo").Value)
        txtGuestOrigin.Text = ConvText(rst("GuestOrigin").Value)
        txtContactNo.Text = ConvText(rst("GuestContact").Value)
        cboTotalGuest.Text = ConvInt(rst("TotalGuest").Value)
        cboStayDuration.Text = ConvInt(rst("StayDuration").Value)
        dtpBookingDate.Value = FormatDate(rst("BookingDate").Value)
        dtpBookingTime.Value = FormatTime(rst("BookingDate").Value)
        dtpCheckInDate.Value = FormatDate(rst("GuestCheckIN").Value)
        dtpCheckInTime.Value = FormatTime(rst("GuestCheckIN").Value)
        dtpCheckOutDate.Value = FormatDate(rst("GuestCheckOUT").Value)
        dtpCheckOutTime.Value = FormatTime(rst("GuestCheckOUT").Value)
        'cboStatus.Text = ConvText(rst("Status").Value)
        txtRemarks.Text = ConvText(rst("Remarks").Value)
        'Room Details
        'lblRoomID.Caption = ConvInt(rst("RoomID").Value)
        lblRoomNo.Caption = ConvText(rst("RoomNo").Value)
'        lblRoomStatus.Caption = rst("RoomStatus").Value
        lblRoomType.Caption = ConvText(rst("RoomType").Value)
        lblLocation.Caption = ConvText(rst("RoomLocation").Value)
        If rst("RoomLocation").Value = True Then
            lblBreakfast.Caption = "Yes"
        Else
            lblBreakfast.Caption = "No"
        End If
        lblBreakfastPrice.Caption = FormatCurrency(rst("BreakfastPrice").Value)
        lblRoomPrice.Caption = FormatCurrency(rst("RoomPrice").Value)
        lblSubTotal.Caption = FormatCurrency(rst("SubTotal").Value)
        txtDeposit.Text = FormatCurrency(rst("Deposit").Value)
        If lblSubTotal.Caption <> "" Then
            dblSubTotal = ConvDbl(lblSubTotal.Caption)
        Else
            dblSubTotal = 0
        End If
        If txtDeposit.Text <> "" Then
            dblDeposit = ConvDbl(txtDeposit.Text)
        Else
            dblDeposit = 0
        End If
        lblTotalDue.Caption = FormatCurrency(dblSubTotal + dblDeposit)
        txtPayment.Text = FormatCurrency(rst("Payment").Value)
        txtRefund.Text = FormatCurrency(rst("Refund").Value)
        If rst("Active").Value = True Then
            strStatus = Trim(rst("RoomStatus").Value)
            ShowStatus strStatus
            tbrMenu.Buttons("VOID").Caption = "Void"
            tbrMenu.Buttons("VOID").ToolTipText = "Void (Ctrl+D)"
            tbrMenu.Buttons("VOID").Image = 5
        Else
            strStatus = "Void"
            ShowStatus strStatus
            tbrMenu.Buttons("VOID").Caption = "Unvoid"
            tbrMenu.Buttons("VOID").ToolTipText = "Unvoid (Ctrl+D)"
            tbrMenu.Buttons("VOID").Image = 6
        End If
        ' Record Details
        'lblCreatedDate.Caption = FormatDateAndTime(rst("CreatedDate").Value)
        'lblCreatedBy.Caption = rst("CreatedBy").Value
        'If rst("LastModifiedDate").Value <> "" Then
        '    lblLastModifiedDate.Caption = FormatDateAndTime(rst("LastModifiedDate").Value)
        'Else
        '    lblLastModifiedDate.Caption = "<Unknown>"
        'End If
        'If rst("LastModifiedBy").Value <> "" Then
        '    lblLastModifiedBy.Caption = rst("LastModifiedBy").Value
        'Else
        '    lblLastModifiedBy.Caption = "<Unknown>"
        'End If
        tbrMenu.Buttons("TEMPORARY").Enabled = True
        tbrMenu.Buttons("OFFICIAL").Enabled = True
        ' Disable Check-IN & Check-OUT buttons
        'tbrMenu.Buttons("Check-IN").Enabled = True
        'tbrMenu.Buttons("Check-OUT").Enabled = True
    Else
        lblBookingID.Caption = Format(lngBookingID, "100000")
        GetRoomDetails intRoomID
        ' Disable Check-IN & Check-OUT buttons
        tbrMenu.Buttons("TEMPORARY").Enabled = False
        tbrMenu.Buttons("OFFICIAL").Enabled = False
        tbrMenu.Buttons("VOID").Enabled = False
        tbrMenu.Buttons("Check-IN").Enabled = False
        tbrMenu.Buttons("Check-OUT").Enabled = False
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

Public Function IsPaid(plngBookingID As Long) As Boolean
    Const mstrMethod As String = "IsPaid"
    Dim rst As ADODB.Recordset
    Dim dblDeposit As Double
    Dim dblPayment As Double
    Dim dblSubTotal As Double
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "SubTotal"
    SQLText "Deposit"
    SQLText "Payment", False
    SQL_FROM "Booking"
    SQL_WHERE_Long "ID", plngBookingID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        dblSubTotal = rst("SubTotal").Value
        dblDeposit = rst("Deposit").Value
        dblPayment = rst("Payment").Value
        If dblPayment = dblSubTotal + dblDeposit Then
            IsPaid = True
        Else
            IsPaid = False
        End If
    Else
        IsPaid = False
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Private Function GetBookingID(intRoomID As Integer) As Long
    Const mstrMethod As String = "GetBookingID"
    Dim rst As ADODB.Recordset
    Dim mlngBookingID As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "ID", False
    SQL_FROM "Booking"
    SQL_WHERE_Integer "RoomID", intRoomID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("ID").Value <> "" Then ' Not Null
            mlngBookingID = ConvLng(rst("ID").Value)
        Else
            mlngBookingID = 0
        End If
    Else
        mlngBookingID = 0
    End If
    GetBookingID = mlngBookingID
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Private Sub GetRoomDetails(intRoomID As Integer)
    Const mstrMethod As String = "GetRoomDetails"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "ID"
    SQLText "BookingID"
    SQLText "RoomStatus"
    SQLText "RoomShortName"
    SQLText "RoomLongName"
    SQLText "RoomType"
    SQLText "RoomLocation"
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
        lblRoomID.Caption = rst("ID").Value
        lngBookingID = rst("BookingID").Value
        lblRoomNo.Caption = rst("RoomShortName").Value
        strStatus = rst("RoomStatus").Value
        ShowStatus strStatus
        lblRoomType.Caption = rst("RoomType").Value
        lblLocation.Caption = rst("RoomLocation").Value
        lblRoomPrice.Caption = FormatCurrency(rst("RoomPrice").Value)
        SumTotal
        lblBreakfastPrice.Caption = FormatCurrency(rst("BreakfastPrice").Value)
        If rst("Breakfast").Value = True Then
            lblBreakfast.Caption = "Yes"
        Else
            lblBreakfast.Caption = "No"
        End If
        '======= Record Details
        'lblCreatedDate.Caption = FormatDateAndTime(rst("CreatedDate").Value)
        'lblCreatedBy.Caption = rst("CreatedBy").Value
        'If rst("LastModifiedDate").Value <> "" Then
        '    lblLastModifiedDate.Caption = FormatDateAndTime(rst("LastModifiedDate").Value)
        'Else
        '    lblLastModifiedDate.Caption = "<Unknown>"
        'End If
        'If rst("LastModifiedBy").Value <> "" Then
        '    lblLastModifiedBy.Caption = rst("LastModifiedBy").Value
        'Else
        '    lblLastModifiedBy.Caption = "<Unknown>"
        'End If
    Else
        ' Room not yet set
        lblRoomID.Caption = intRoomID
        lngBookingID = 0
        lblRoomNo.Caption = ""
        'lblRoomStatus.Caption = ""
        'lblRoomStatus.BackColor = &HFF00&
        lblRoomType.Caption = ""
        lblLocation.Caption = ""
        lblRoomPrice.Caption = "0.00"
        lblBreakfastPrice.Caption = "0.00"
        lblBreakfast.Caption = ""
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

Public Sub SelectRoom(pintRoomID As Integer)
    Const mstrMethod As String = "SelectRoom"
On Error GoTo CheckErr
    intRoomID = pintRoomID
    lblRoomID.Caption = intRoomID
    GetRoomDetails intRoomID
    If lngBookingID = 0 Then
        'ResetFields
        'strStatus = "Open"
        'ShowStatus strStatus
        CreateTempBookingID
        lblBookingID.Caption = Format(lngBookingID, "100000")
        ' Disable Check-IN & Check-OUT buttons
        tbrMenu.Buttons("TEMPORARY").Enabled = False
        tbrMenu.Buttons("OFFICIAL").Enabled = False
        tbrMenu.Buttons("VOID").Enabled = False
        tbrMenu.Buttons("Check-IN").Enabled = False
        tbrMenu.Buttons("Check-OUT").Enabled = False
    Else
        PopulateValues lngBookingID
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub CreateTempBookingID()
    Const mstrMethod As String = "CreateTempBookingID"
On Error GoTo CheckErr
    lngBookingID = GetTempBookingID
    If lngBookingID = 0 Then
        'INSERT New Sales Record
        SQL_INSERT "Booking"
        SQLText "CreatedDate"
        SQLText "CreatedBy"
        SQLText "Active"
        SQLText "Temp", False
        SQL_VALUES
        SQLData_DateTime FormatDateAndTime(Now)
        SQLData_Text gstrUserID
        SQLData_Boolean True
        SQLData_Boolean True, False
        SQL_Close_Bracket
        OpenDB
        QuerySQL gstrSQL
        CloseDB
        lngBookingID = GetTempBookingID
    End If
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Function GetTempBookingID() As Long
    Const mstrMethod As String = "GetTempBookingID"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    OpenDB
    SQL_SELECT_TOP "ID", "Booking"
    SQL_WHERE_Boolean "Temp", True
    SQL_ORDER_BY "ID", False
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        GetTempBookingID = ConvLng(rst("ID").Value)
    Else
        GetTempBookingID = 0
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Private Sub SaveBooking()
    Const mstrMethod As String = "SaveBooking"
    Dim rst As ADODB.Recordset
    Dim lngCheckingBookingID As Long
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    If Trim(txtGuestName.Text) = "" Then
        MsgBox "Please key in Guest Name", vbExclamation, mstrMethod
        txtGuestName.SetFocus
        Exit Sub
    End If
    If Trim(txtGuestPassport.Text) = "" Then
        MsgBox "Please key in Guest Passport/IC No", vbExclamation, mstrMethod
        txtGuestPassport.SetFocus
        Exit Sub
    End If
    If cboTotalGuest.ListIndex < 0 Then
        MsgBox "Please select Total Guest", vbExclamation, mstrMethod
        cboTotalGuest.SetFocus
        Exit Sub
    End If
    If cboStayDuration.ListIndex < 0 Then
        MsgBox "Please select Stay Duration", vbExclamation, mstrMethod
        cboStayDuration.SetFocus
        Exit Sub
    End If
    ' ====== Validate Check-In Date/Time ======
    ' Check OUT cannot smaller/ealier than Check IN
    
    ' ====== Validate Room Details ======
    ' Do not leave empty
    
    If vbNo = MsgBox("Do you want to Save this Booking?", vbQuestion + vbYesNo, mstrMethod) Then
        Exit Sub
    End If
    
    'Update table Booking
    SQL_UPDATE "Booking"
    SQL_SET_Text "GuestName", Trim(txtGuestName.Text)
    SQL_SET_Text "GuestPassport", Trim(txtGuestPassport.Text)
    SQL_SET_Text "GuestEmergencyContactName", Trim(txtGuestEmergencyContactName.Text)
    SQL_SET_Text "GuestEmergencyContactNo", Trim(txtGuestEmergencyContactNo.Text)
    SQL_SET_Text "GuestOrigin", Trim(txtGuestOrigin.Text)
    SQL_SET_Text "GuestContact", Trim(txtContactNo.Text)
    SQL_SET_DateTime "BookingDate", FormatDate(dtpBookingDate.Value) & " " & FormatTime(dtpBookingTime.Value)
    SQL_SET_DateTime "GuestCheckIN", FormatDate(dtpCheckInDate.Value) & " " & FormatTime(dtpCheckInTime.Value)
    SQL_SET_DateTime "GuestCheckOUT", FormatDate(dtpCheckOutDate.Value) & " " & FormatTime(dtpCheckOutTime.Value)
    SQL_SET_Integer "TotalGuest", CInt(cboTotalGuest.Text)
    SQL_SET_Integer "StayDuration", CInt(cboStayDuration.Text)
    'SQL_SET_Text "RoomStatus", Trim(cboStatus.Text)
    SQL_SET_Text "Remarks", Trim(txtRemarks.Text)
    ' Room Details
    SQL_SET_Integer "RoomID", Trim(lblRoomID.Caption)
    SQL_SET_Text "RoomNo", Trim(lblRoomNo.Caption)
    SQL_SET_Text "RoomType", Trim(lblRoomType.Caption)
    SQL_SET_Text "RoomLocation", Trim(lblLocation.Caption)
    SQL_SET_Double "RoomPrice", Val(lblRoomPrice.Caption)
    If lblBreakfast.Caption = "Yes" Then
        SQL_SET_Boolean "Breakfast", True
    Else
        SQL_SET_Boolean "Breakfast", False
    End If
    SQL_SET_Double "BreakfastPrice", Val(lblBreakfastPrice.Caption)
    SQL_SET_Double "SubTotal", Val(lblSubTotal.Caption)
    SQL_SET_Double "Deposit", Val(txtDeposit.Text)
    SQL_SET_Double "Payment", Val(txtPayment.Text)
    SQL_SET_Double "Refund", Val(txtRefund.Text)
    ' Record Details
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID
    If strStatus = "Open" Then
        strStatus = "Booked"
        SQL_SET_DateTime "CreatedDate", FormatDateAndTime(Now)
        SQL_SET_Text "CreatedBy", gstrUserID
    End If
    SQL_SET_Boolean "Temp", False, False
    SQL_WHERE_Long "ID", lngBookingID
    'CloseRS rst
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    'ResetFields
    UpdateRoomStatus strStatus, intRoomID, lngBookingID
    ShowStatus strStatus
    MsgBox "Booking is updated!", vbInformation, mstrMethod
    ' Note: If Closed then Booking ID =  0 and cannot print
    'intRoomID = ConvInt(lblRoomID.Caption)
    'MsgBox "Booking is saved!", vbInformation, mstrMethod
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Check_IN()
    Const mstrMethod As String = "Check-IN"
    Dim rst As ADODB.Recordset
    Dim intDay As Integer
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    If lngBookingID = 0 Then
        MsgBox "Booking has not yet saved. Please save it first!", vbExclamation, mstrMethod
        Exit Sub
    End If
    If vbNo = MsgBox("Do you want to Check-IN this Room at " & FormatDate(dtpCheckInDate.Value) & " " & FormatTime(dtpCheckInTime.Value) & " ?", vbQuestion + vbYesNo, mstrMethod) Then
        Exit Sub
    End If
    'Check if full payment is made
    If IsPaid(lngBookingID) = False Then
        MsgBox "Please make payment first!", vbInformation, mstrMethod
        Exit Sub
    End If
    'Update detals in Booking table
    SQL_UPDATE "Booking"
    SQL_SET_DateTime "GuestCheckIN", FormatDate(dtpCheckInDate.Value) & " " & FormatTime(dtpCheckInTime.Value)
    'SQL_SET_DateTime "GuestCheckOUT", FormatDate(dtpCheckOutDate.Value) & " " & FormatTime(dtpCheckOutTime.Value)
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Long "ID", lngBookingID
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    ' Change Room Status to OCCUPIED
    strStatus = "Occupied"
    UpdateRoomStatus strStatus, intRoomID
    MsgBox "Room is Checked In!", vbInformation, mstrMethod '& " (" & lngRecordsAffected & ")"
    PopulateValues lngBookingID
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Check_OUT()
    Const mstrMethod As String = "Check-OUT"
    Dim rst As ADODB.Recordset
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    If lngBookingID = 0 Then
        MsgBox "Booking has not yet saved. Please save it first!", vbExclamation, mstrMethod
        Exit Sub
    End If
    'Check if full payment is made
    If IsPaid(lngBookingID) = False Then
        MsgBox "Please make payment first!", vbInformation, mstrMethod
        Exit Sub
    End If
    ' If Late Check-OUT then Deposit is 'NO REFUND'
    If TimeValue(dtpCheckOutTime.Value) >= "2:00 PM" Then
        MsgBox "REFUND = MYR 0.00" & vbCrLf & _
        "Deposit Refund is not allowed for Check Out after 2:00 PM!", vbExclamation, "DEPOSIT NO REFUND"
        txtRefund.Text = "0.00"
        txtRefund.Enabled = False
    Else
        txtRefund.Enabled = True
        If vbYes = MsgBox("Do you want to fully REFUND the Deposit?" & vbCrLf & vbCrLf & _
            "Select 'Yes' for REFUND = MYR " & Trim(txtDeposit.Text) & vbCrLf & _
            "Select 'No'  for REFUND = MYR " & Trim(txtRefund.Text), vbQuestion + vbYesNo, mstrMethod) Then
            txtRefund.Text = Trim(txtDeposit.Text)
        End If
    End If
    If vbNo = MsgBox("REFUND = MYR " & FormatCurrency(txtRefund.Text) & vbCrLf & _
        "Do you want to Check-OUT this Room at " & _
        FormatDate(dtpCheckOutDate.Value) & " " & _
        FormatTime(dtpCheckOutTime.Value) & " ?", vbQuestion + vbYesNo, mstrMethod) Then
        Exit Sub
    End If
    'Update detals in Booking table
    SQL_UPDATE "Booking"
    SQL_SET_DateTime "GuestCheckOUT", FormatDate(dtpCheckOutDate.Value) & " " & FormatTime(dtpCheckOutTime.Value) ' FormatDateAndTime(Now)
    ' Update Refund
    SQL_SET_Double "Refund", FormatCurrency(txtRefund.Text)
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Long "ID", lngBookingID
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    ' Change Room Status to HOUSEKEEPING
    strStatus = "Housekeeping"
    UpdateRoomStatus strStatus, intRoomID
    MsgBox "Room is Checked Out!", vbInformation, mstrMethod '& " (" & lngRecordsAffected & ")"
    PopulateValues lngBookingID
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub UpdateRoomStatus(pstrStatus As String, pintRoomID As Integer, Optional plngBookingID As Long = 0)
    Const mstrMethod As String = "UpdateRoomStatus"
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    If pintRoomID = 0 Then Exit Sub
    SQL_UPDATE "Room"
    SQL_SET_Text "RoomStatus", pstrStatus
    Select Case pstrStatus
    Case "Open"
        ' SQL_SET_Long "BookingID", 0
    Case "Booked"
        SQL_SET_Long "BookingID", plngBookingID
    Case "Occupied"
        ' No update
    Case "Housekeeping"
        'SQL_SET_Long "BookingID", 0
    Case Else ' Maintenance
        ' No update
    End Select
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Integer "ID", pintRoomID
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Sub VoidBooking(plngBookingID As Long)
    Const mstrMethod As String = "VoidBooking"
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    If plngBookingID = 0 Then Exit Sub
    SQL_UPDATE "Booking"
    If strStatus = "Void" Then
        SQL_SET_Boolean "Active", True
    Else
        SQL_SET_Boolean "Active", False
    End If
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Long "ID", plngBookingID
    SQLText "AND Temp = FALSE", False
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    PopulateValues plngBookingID
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub PrintReceipt(strType As String)
    Const mstrMethod As String = "Print Receipt"
    Dim rep As New frmPrint
On Error GoTo CheckErr
    If strType = "OFFICIAL" Then 'If strType = "Closed" Or strType = "Housekeeping" Then
        gstrReportFileName = "Official Receipt.rpt"
        gstrReportTitle = "OFFICIAL RECEIPT"
        gstrSQL = "SELECT C.CompanyName, C.StreetAddress, C.ContactNo,"
        gstrSQL = gstrSQL & " Format(B.ID, '100000') AS BookingID, B.GuestName, B.GuestCheckIN, B.GuestCheckOUT,"
        gstrSQL = gstrSQL & " B.RoomType, B.Payment, B.Refund, B.Payment-B.Refund AS Total, B.CreatedDate, '" & gstrUserID & "' AS IssuedBy" 'B.CreatedBy"
        gstrSQL = gstrSQL & " FROM Company C, Booking B"
        gstrSQL = gstrSQL & " WHERE B.ID = " & lngBookingID
        If QueryHasData(gstrSQL) = False Then
            gstrSQL = "SELECT CompanyName, StreetAddress, ContactNo,"
            gstrSQL = gstrSQL & " '000000' AS BookingID, '' AS GuestName, '' AS GuestCheckIN, '' AS GuestCheckOUT,"
            gstrSQL = gstrSQL & " '' AS RoomType, 0 AS Payment, 0 AS Refund, 0 AS Total, '' AS CreatedDate, '' AS IssuedBy"
            gstrSQL = gstrSQL & " FROM Company"
        End If
        rep.Caption = "OFFICIAL RECEIPT (" & Format(lngBookingID, "100000") & ")"
    Else
        gstrReportFileName = "Temporary Receipt.rpt"
        gstrReportTitle = "TEMPORARY RECEIPT"
        gstrSQL = "SELECT C.CompanyName, C.StreetAddress, C.ContactNo,"
        gstrSQL = gstrSQL & " Format(B.ID, '100000') AS BookingID, B.GuestName, B.GuestCheckIN, B.GuestCheckOUT,"
        gstrSQL = gstrSQL & " B.RoomType, B.Payment-B.Deposit AS SubTotal, B.Deposit, B.Payment AS Total, B.CreatedDate, '" & gstrUserID & "' AS IssuedBy" 'B.CreatedBy"
        gstrSQL = gstrSQL & " FROM Company C, Booking B"
        gstrSQL = gstrSQL & " WHERE B.ID = " & lngBookingID
        If QueryHasData(gstrSQL) = False Then
            gstrSQL = "SELECT CompanyName, StreetAddress, ContactNo,"
            gstrSQL = gstrSQL & " '000000' AS BookingID, '' AS GuestName, '' AS GuestCheckIN, '' AS GuestCheckOUT,"
            gstrSQL = gstrSQL & " '' AS RoomType, 0 AS SubTotal, 0 AS Deposit, 0 AS Total, '' AS CreatedDate, '' AS IssuedBy"
            gstrSQL = gstrSQL & " FROM Company"
        End If
        rep.Caption = "TEMPORARY RECEIPT (" & Format(lngBookingID, "100000") & ")"
    End If
    'Show the Report
    rep.Show
    Exit Sub
CheckErr:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub ResetFields()
    ' Booking Details
    txtGuestName.Text = ""
    txtGuestPassport.Text = ""
    txtGuestEmergencyContactName.Text = ""
    txtGuestEmergencyContactNo.Text = ""
    txtGuestOrigin.Text = ""
    txtContactNo.Text = ""
    txtRemarks.Text = ""
    txtDeposit.Text = "20.00"
    txtPayment.Text = "0.00"
    txtRefund.Text = "0.00"
    dtpBookingDate.Value = FormatDate(Now)
    dtpBookingTime.Value = FormatTime(Now)
    'Check IN
    dtpCheckInDate.Value = FormatDate(Now)
    dtpCheckInTime.Value = FormatTime(Now) 'FormatDate(dtpCheckInDate.Value) & " 12:00:00 PM"
    cboTotalGuest.ListIndex = 0 '- 1
    cboStayDuration.ListIndex = 0 '-1
    SumTotal
End Sub

Private Sub DisableControls()
    dtpBookingDate.Enabled = False
    dtpBookingTime.Enabled = False
    cboTotalGuest.Enabled = False
    cboStayDuration.Enabled = False
    dtpCheckInDate.Enabled = False
    dtpCheckInTime.Enabled = False
    dtpCheckOutDate.Enabled = False
    dtpCheckOutTime.Enabled = False
    txtGuestName.Enabled = False
    txtGuestPassport.Enabled = False
    txtGuestOrigin.Enabled = False
    txtContactNo.Enabled = False
    txtGuestEmergencyContactName.Enabled = False
    txtGuestEmergencyContactNo.Enabled = False
    txtDeposit.Enabled = False
    txtPayment.Enabled = False
    txtRefund.Enabled = False
    txtRemarks.Enabled = False
End Sub

Private Sub txtDeposit_LostFocus()
    Dim dblTotalDue As Double
    dblTotalDue = ConvDbl(lblSubTotal.Caption) + ConvDbl(txtDeposit.Text)
    lblTotalDue.Caption = FormatCurrency(dblTotalDue)
End Sub

Private Sub txtPayment_LostFocus()
    Dim dblPayment As Double
    dblPayment = ConvDbl(txtPayment.Text)
    txtPayment.Text = FormatCurrency(dblPayment)
End Sub

Private Sub txtRefund_LostFocus()
    Dim dblRefund As Double
    dblRefund = ConvDbl(txtRefund.Text)
    txtRefund.Text = FormatCurrency(dblRefund)
End Sub
