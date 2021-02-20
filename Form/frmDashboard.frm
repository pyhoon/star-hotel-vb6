VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDashboard 
   BackColor       =   &H00505050&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15375
   Icon            =   "frmDashboard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1455
      ScaleWidth      =   14895
      TabIndex        =   75
      Top             =   2640
      Width           =   14895
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "052"
         Height          =   975
         Index           =   52
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "053"
         Height          =   975
         Index           =   53
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "054"
         Height          =   975
         Index           =   54
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "055"
         Height          =   975
         Index           =   55
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "051"
         Height          =   975
         Index           =   51
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "050"
         Height          =   975
         Index           =   50
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "049"
         Height          =   975
         Index           =   49
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "048"
         Height          =   975
         Index           =   48
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "047"
         Height          =   975
         Index           =   47
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "046"
         Height          =   975
         Index           =   46
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "045"
         Height          =   975
         Index           =   45
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   320
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Level 4"
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
         Height          =   300
         Left            =   240
         TabIndex        =   76
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1455
      ScaleWidth      =   14895
      TabIndex        =   73
      Top             =   4080
      Width           =   14895
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "040"
         Height          =   975
         Index           =   40
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "041"
         Height          =   975
         Index           =   41
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "042"
         Height          =   975
         Index           =   42
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "043"
         Height          =   975
         Index           =   43
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "044"
         Height          =   975
         Index           =   44
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "039"
         Height          =   975
         Index           =   39
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "038"
         Height          =   975
         Index           =   38
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "037"
         Height          =   975
         Index           =   37
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "036"
         Height          =   975
         Index           =   36
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "035"
         Height          =   975
         Index           =   35
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   320
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "034"
         Height          =   975
         Index           =   34
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   320
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Level 3"
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
         Height          =   300
         Left            =   240
         TabIndex        =   74
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   240
      ScaleHeight     =   2520
      ScaleWidth      =   14895
      TabIndex        =   71
      Top             =   5520
      Width           =   14895
      Begin VB.CommandButton cmdUnit 
         Caption         =   "012"
         Height          =   975
         Index           =   12
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "020"
         Height          =   975
         Index           =   20
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "018"
         Height          =   975
         Index           =   18
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "016"
         Height          =   975
         Index           =   16
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "014"
         Height          =   975
         Index           =   14
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "013"
         Height          =   975
         Index           =   13
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "015"
         Height          =   975
         Index           =   15
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "017"
         Height          =   975
         Index           =   17
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "019"
         Height          =   975
         Index           =   19
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "021"
         Height          =   975
         Index           =   21
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "023"
         Height          =   975
         Index           =   23
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "024"
         Height          =   975
         Index           =   24
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "025"
         Height          =   975
         Index           =   25
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "026"
         Height          =   975
         Index           =   26
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Appearance      =   0  'Flat
         Caption         =   "027"
         Height          =   975
         Index           =   27
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "033"
         Height          =   975
         Index           =   33
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "032"
         Height          =   975
         Index           =   32
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "031"
         Height          =   975
         Index           =   31
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "030"
         Height          =   975
         Index           =   30
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "029"
         Height          =   975
         Index           =   29
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "028"
         Height          =   975
         Index           =   28
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "022"
         Height          =   975
         Index           =   22
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Level 2"
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
         Height          =   300
         Left            =   240
         TabIndex        =   72
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1455
      ScaleWidth      =   14895
      TabIndex        =   69
      Top             =   8040
      Width           =   14895
      Begin VB.CommandButton cmdUnit 
         Caption         =   "001"
         Height          =   975
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "002"
         Height          =   975
         Index           =   2
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "003"
         Height          =   975
         Index           =   3
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "004"
         Height          =   975
         Index           =   4
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "005"
         Height          =   975
         Index           =   5
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "006"
         Height          =   975
         Index           =   6
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "007"
         Height          =   975
         Index           =   7
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "008"
         Height          =   975
         Index           =   8
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "009"
         Height          =   975
         Index           =   9
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "010"
         Height          =   975
         Index           =   10
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   315
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "011"
         Height          =   975
         Index           =   11
         Left            =   13440
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Level 1"
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
         Height          =   300
         Left            =   240
         TabIndex        =   70
         Top             =   0
         Width           =   1920
      End
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Height          =   810
      Left            =   0
      TabIndex        =   56
      Top             =   1200
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   1429
      ButtonWidth     =   2011
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
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
            Caption         =   "Report (F2)"
            Key             =   "REPORT"
            Object.ToolTipText     =   "Open Report (F2)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Customer (F3)"
            Key             =   "CUSTOMER"
            Object.ToolTipText     =   "Find Customer (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Room (F4)"
            Key             =   "ROOM"
            Object.ToolTipText     =   "Maintain Room (F4)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User (F5)"
            Key             =   "USER"
            Object.ToolTipText     =   "Maintain User (F5)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Access (F6)"
            Key             =   "ACCESS"
            Object.ToolTipText     =   "Access Control (F6)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Blink (F7)"
            Key             =   "BLINK"
            Object.ToolTipText     =   "Blink Button (F7)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Security (F8)"
            Key             =   "SECURITY"
            Object.ToolTipText     =   "Change Password (F8)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmDashboard.frx":0CCA
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   64
         Top             =   0
         Width           =   4935
         Begin VB.Timer tmrBlink 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   480
            Top             =   120
         End
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
            TabIndex        =   66
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
            TabIndex        =   65
            Top             =   390
            Width           =   1620
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9720
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   -2147483630
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":0FE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":18BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":2198
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":305A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":3934
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":47F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":54D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":5922
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDashboard.frx":5D74
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdUnit 
      Caption         =   "R000"
      Height          =   495
      Index           =   0
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   57
      Top             =   1900
      Width           =   15455
      Begin VB.Shape Shape5 
         BackColor       =   &H00FF7929&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   12720
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblSummary5 
         BackStyle       =   0  'Transparent
         Caption         =   "Maintenance : 0"
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
         Height          =   240
         Left            =   13080
         TabIndex        =   63
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label lblSummary4 
         BackStyle       =   0  'Transparent
         Caption         =   "Housekeeping : 0"
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
         Height          =   240
         Left            =   10080
         TabIndex        =   62
         Top             =   240
         Width           =   1980
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00F900D5&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   9720
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblSummary3 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupied : 0"
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
         Height          =   240
         Left            =   7560
         TabIndex        =   61
         Top             =   240
         Width           =   1980
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H004417FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   7200
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblSummary2 
         BackStyle       =   0  'Transparent
         Caption         =   "Booked  : 0"
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
         Height          =   240
         Left            =   5040
         TabIndex        =   60
         Top             =   240
         Width           =   1980
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0000EAFF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4680
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblSummary1 
         BackStyle       =   0  'Transparent
         Caption         =   "Open : 0"
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
         Height          =   240
         Left            =   2520
         TabIndex        =   59
         Top             =   240
         Width           =   1980
      End
      Begin VB.Label lblSummary 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Summary"
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
         Height          =   240
         Left            =   360
         TabIndex        =   58
         Top             =   240
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0076E600&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   2160
         Top             =   240
         Width           =   255
      End
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
      TabIndex        =   68
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
      TabIndex        =   67
      Top             =   900
      Width           =   5295
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmDashboard.frx":7AC6
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
   Begin VB.Menu mnuPop 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnuBooking 
         Caption         =   "Booking"
      End
      Begin VB.Menu mnuEditRoom 
         Caption         =   "Edit Room"
      End
      Begin VB.Menu mnuChangeStatus 
         Caption         =   "Change Status"
         Begin VB.Menu mnuFree 
            Caption         =   "Free"
         End
         Begin VB.Menu mnuOccupied 
            Caption         =   "Occupied"
         End
         Begin VB.Menu mnuHousekeeping 
            Caption         =   "Housekeeping"
         End
         Begin VB.Menu mnuMaintenance 
            Caption         =   "Maintenance"
         End
      End
   End
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
'
' Modified On : 02/11/2014
' Descriptions : 1) Do not blink Room where status = Maintenance
'
' Modified On : 28/10/2014
' Descriptions : 1) Add function NeedBlink to check and enable timer if return value is TRUE

' Nice Theme: http://keenthemes.com/preview/metronic/theme/admin_2/ui_colors.html
' Material Design: https://material.io/design/color/the-color-system.html#tools-for-picking-colors
' VB6 Control Backcolor is reversed RGB Hex
' Example: Blue-Soft #4C87C9
' We separate the code to 3 parts
' R=4C, G=87, B=C9
' VB Control colour code
' &H00  C9  87  4C  &
'       B   G   R
' So the code is &H00C9874C& (13207372)

Option Explicit
Private Const mstrModule As String = "Dashboard"
Private Const COL_YELLOW = &HEAFF&
Private Const COL_GREEN = &H76E600
Private Const COL_BLUE = &HFF7929
Private Const COL_RED = &H4417FF
Private Const COL_PURPLE = &HF900D5
Private Const COL_GRAY = &H505050
Dim blnBlink(61) As Boolean
Dim strBackColor(61) As String
Dim blnDim As Boolean
Dim intButtonIndex As Integer
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
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    tmrClock.Enabled = True
    intTick = 0
End Sub

Private Sub Form_Deactivate()
    tmrClock.Enabled = False
End Sub

Private Sub Form_Load()
    lblBusinessName.Caption = COMPANY_PRODUCT_NAME
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    ReCaptionButton
    SetButtonProperties
    LoadBlinkSetting
    ShowSummary1
    ShowSummary2
    ShowSummary3
    ShowSummary4
    ShowSummary5
    If UserAccessModule(MOD_REPORT_LIST) = True Then
        tbrMenu.Buttons("REPORT").Enabled = True
    Else
        tbrMenu.Buttons("REPORT").Enabled = False
    End If
    If UserAccessModule(MOD_FIND_CUSTOMER) = True Then
        tbrMenu.Buttons("CUSTOMER").Enabled = True
    Else
        tbrMenu.Buttons("CUSTOMER").Enabled = False
    End If
    If UserAccessModule(MOD_MAINTAIN_ROOM) = True Then
        tbrMenu.Buttons("ROOM").Enabled = True
        mnuEditRoom.Enabled = True
    Else
        tbrMenu.Buttons("ROOM").Enabled = False
        mnuEditRoom.Enabled = False
    End If
    If UserAccessModule(MOD_MAINTAIN_USER) = True Then
        tbrMenu.Buttons("USER").Enabled = True
    Else
        tbrMenu.Buttons("USER").Enabled = False
    End If
    If UserAccessModule(MOD_ACCESS_CONTROL) = True Then
        tbrMenu.Buttons("ACCESS").Enabled = True
    Else
        tbrMenu.Buttons("ACCESS").Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmUserLogin.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyF2 And tbrMenu.Buttons("REPORT").Enabled Then
        frmReport.Show
        Me.Hide
    ElseIf KeyCode = vbKeyF3 And tbrMenu.Buttons("CUSTOMER").Enabled Then
        frmFindCustomer.Show
        Me.Hide
    ElseIf KeyCode = vbKeyF4 And tbrMenu.Buttons("ROOM").Enabled Then
        'frmRoomMaintain.SelectRoom 1
        frmRoomMaintain.Show
        Me.Hide
    ElseIf KeyCode = vbKeyF5 And tbrMenu.Buttons("USER").Enabled Then
        frmUserMaintain.Show
        Me.Hide
    ElseIf KeyCode = vbKeyF6 And tbrMenu.Buttons("ACCESS").Enabled Then
        frmModuleAccess.Show
        Me.Hide
    ElseIf KeyCode = vbKeyF7 And tbrMenu.Buttons("BLINK").Enabled Then
        BlinkButton
    ElseIf KeyCode = vbKeyF8 And tbrMenu.Buttons("SECURITY").Enabled Then
        gblnUserChangePassword = True
        frmUserChangePassword.Show
        Me.Hide
    Else ' KeyCode = vbKeyB And (Shift And vbCtrlMask) And tbrMenu.Buttons("BLINK").Enabled Then
        Exit Sub
    End If
End Sub

Private Sub mnuBooking_Click()
    cmdUnit_Click intButtonIndex
End Sub

Private Sub mnuEditRoom_Click()
    With frmRoomMaintain
        .SelectRoom intButtonIndex
        '.PopulateValues intButtonIndex
        '.lvRooms_Click
        .Show
    End With
End Sub

Private Sub mnuFree_Click()
    Const mstrMethod As String = "mnuFree"
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    'Update detals in Room table
    SQL_UPDATE "Room"
    'SQL_SET_Long "RoomID", 0
    'SQL_SET_Boolean "Maintenance", False
    SQL_SET_Text "RoomStatus", "Open"
    SQL_SET_Long "BookingID", 0
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Integer "ID", intButtonIndex
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
'    'Update detals in Room table
'    SQL_UPDATE "Booking"
'    SQL_SET_Long "RoomID", 0
'    SQL_SET_Boolean "Maintenance", False
'    SQL_SET_Text "Status", "Open"
'    ' === Direct update LastModifiedDate ??? ===
'    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
'    SQL_SET_Text "LastModifiedBy", gstrUserID, False
'    SQL_WHERE_Integer "RoomID", intButtonIndex
'    OpenDB
'    QuerySQL gstrSQL, lngRecordsAffected
'    CloseDB
    ShowSummary1
    'ShowSummary2
    'ShowSummary3
    ShowSummary4
    ShowSummary5
    SetButtonProperties
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub mnuOccupied_Click()
    Const mstrMethod As String = "mnuOccupied"
    Dim mlngBookingID As Long
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    If vbNo = MsgBox("Do you want to Check-IN this Room at " & FormatDateAndTime(Now) & " ?", vbQuestion + vbYesNo, "Check-IN") Then
        Exit Sub
    End If
    mlngBookingID = GetBookingID(intButtonIndex)
    
    'Update detals in Booking table
    SQL_UPDATE "Booking"
    SQL_SET_DateTime "GuestCheckIN", FormatDateAndTime(Now)
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Long "ID", mlngBookingID
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    
    'Update detals in Room table
    SQL_UPDATE "Room"
    SQL_SET_Text "RoomStatus", "Occupied"
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Integer "ID", intButtonIndex
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    MsgBox "Booking is updated!", vbInformation, mstrMethod
    'ShowSummary1
    ShowSummary2
    ShowSummary3
    'ShowSummary4
    'ShowSummary5
    SetButtonProperties
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub mnuHousekeeping_Click()
    Const mstrMethod As String = "mnuHousekeeping"
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    'Update detals in Room table
    SQL_UPDATE "Room"
    SQL_SET_Text "RoomStatus", "Housekeeping"
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Integer "ID", intButtonIndex
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    ShowSummary1
    'ShowSummary2
    'ShowSummary3
    ShowSummary4
    ShowSummary5
    SetButtonProperties
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub mnuMaintenance_Click()
    Const mstrMethod As String = "mnuMaintenance"
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    'Update detals in Room table
    SQL_UPDATE "Room"
    'SQL_SET_Long "BookingID", lngBookingID
    'If blnSetHousekeeping = True Then
        'SQL_SET_Boolean "Maintenance", True
        SQL_SET_Text "RoomStatus", "Maintenance", True
        'SQL_SET_Text "Status", "Booked"
    'End If
    ' === Direct update LastModifiedDate ??? ===
    SQL_SET_DateTime "LastModifiedDate", FormatDateAndTime(Now)
    SQL_SET_Text "LastModifiedBy", gstrUserID, False
    SQL_WHERE_Integer "ID", intButtonIndex
    OpenDB
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    ShowSummary1
    'ShowSummary2
    'ShowSummary3
    ShowSummary4
    ShowSummary5
    SetButtonProperties
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub cmdUnit_Click(Index As Integer)
    Const mstrMethod As String = "cmdUnit_Click"
On Error GoTo CheckErr
    'Beep
    'PlayWav "Chimes.wav"
    'frmGuestDetails.Show
    'frmRoomMaintain.PopulateValues CLng(Index)
    'frmRoomMaintain.lvRooms.ListItems(Index).Selected = True
        
    If UserAccessModule(MOD_BOOKING) = False Then
        MsgBox "You do not have access to make booking!", vbExclamation, "Access Denied"
        Exit Sub
    End If

    ' Check if Room is under Maintenance
    If GetRoomStatus(Index) = "Maintenance" Then
        MsgBox "Room is under Maintenance. Please choose other Room.", vbExclamation, "Room Maintenance"
        Exit Sub
    End If
'    ' Check if Room is under Housekeeping
'    If GetRoomStatus(Index) = "Housekeeping" Then
'        MsgBox "Room is under Housekeeping. Please choose other Room.", vbExclamation, "Room Housekeeping"
'        Exit Sub
'    End If
    
    ' Check if Room is exist in table Room
    If RoomSetup(Index) = False Then
        'MsgBox "Room is not setup yet. Please setup using Room Maintenance.", vbExclamation, "Room Setup"
        ' Make sure User has Module Access !
        If UserAccessModule(MOD_MAINTAIN_ROOM) = True Then
            If vbYes = MsgBox("Room is not setup yet. Do you want to Setup now?", vbQuestion + vbYesNo, "Room Setup") Then
                With frmRoomMaintain
                    .PopulateValues Index
                    .Show
                End With
            End If
        End If
        Exit Sub
    End If
       
    With frmBooking
        .SelectRoom Index
        .Show
    End With
    Me.Hide
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub cmdUnit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    intButtonIndex = Index
    If Button = 2 Then
        Select Case GetRoomStatus(Index)
        Case "Open"
            mnuBooking.Enabled = True
            mnuChangeStatus.Enabled = True
            mnuFree.Enabled = False
            mnuOccupied.Enabled = False
            mnuHousekeeping.Enabled = True
            mnuMaintenance.Enabled = True
            PopupMenu mnuPop, , , , mnuBooking
        Case "Booked"
            mnuBooking.Enabled = True
            mnuChangeStatus.Enabled = True
            mnuFree.Enabled = False
            mnuOccupied.Enabled = True
            mnuHousekeeping.Enabled = False
            mnuMaintenance.Enabled = False
            PopupMenu mnuPop, , , , mnuBooking
        Case "Occupied"
            mnuBooking.Enabled = True
            mnuChangeStatus.Enabled = False
            'mnuFree.Enabled = False
            'mnuOccupied.Enabled = False
            PopupMenu mnuPop, , , , mnuBooking
        Case "Housekeeping"
            mnuBooking.Enabled = False
            mnuChangeStatus.Enabled = True
            mnuFree.Enabled = True
            mnuOccupied.Enabled = False
            mnuHousekeeping.Enabled = False
            mnuMaintenance.Enabled = True
            PopupMenu mnuPop, , , , mnuBooking
        Case "Maintenance"
            mnuBooking.Enabled = False
            mnuChangeStatus.Enabled = True
            mnuFree.Enabled = True
            mnuOccupied.Enabled = False
            mnuHousekeeping.Enabled = False
            mnuMaintenance.Enabled = False
            PopupMenu mnuPop, , , , mnuBooking
        Case Else
        
        End Select
    End If
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "CLOSE"
            Unload Me
        Case "REPORT"
            frmReport.Show
            Me.Hide
        Case "CUSTOMER"
            frmFindCustomer.Show
            Me.Hide
        Case "ROOM"
            'frmRoomMaintain.SelectRoom 1
            frmRoomMaintain.Show
            Me.Hide
        Case "USER"
            frmUserMaintain.Show
            Me.Hide
        Case "ACCESS"
            frmModuleAccess.Show
            Me.Hide
        Case "BLINK"
            BlinkButton
        Case "SECURITY"
            gblnUserChangePassword = True
            frmUserChangePassword.Show
            Me.Hide
        Case Else 'Default
            Exit Sub
    End Select
End Sub

Private Sub tmrBlink_Timer()
    Dim i As Integer
    If blnDim Then
        For i = 1 To 61
            If blnBlink(i) = True Then
                'cmdUnit(i).BackColor = &H505050    ' &HC0C0C0   '&H80000015
                If strBackColor(i) = COL_YELLOW Then
                    cmdUnit(i).BackColor = COL_YELLOW
                End If
                If strBackColor(i) = COL_RED Then
                    cmdUnit(i).BackColor = COL_GRAY
                End If
            End If
        Next
    Else
        For i = 1 To 61
            If blnBlink(i) = True Then
                'cmdUnit(i).BackColor = strBackColor(i)
                If strBackColor(i) = COL_YELLOW Then
                    cmdUnit(i).BackColor = COL_GRAY
                End If
                If strBackColor(i) = COL_RED Then
                    cmdUnit(i).BackColor = COL_RED
                End If
            End If
        Next
    End If
    blnDim = Not blnDim
End Sub

Private Sub ReCaptionButton()
    Dim i As Integer
    For i = 1 To cmdUnit.UBound
        cmdUnit(i).Caption = "" '"<Not Set>"
        cmdUnit(i).BackColor = COL_GRAY ' &HC0C0C0
    Next
End Sub

Public Sub ShowSummary5()
    Const mstrMethod As String = "ShowSummary5"
    Dim rst As ADODB.Recordset
    Dim intSum5 As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "COUNT(ID) AS SUM5", False
    SQL_FROM "Room"
    SQL_WHERE_Text "RoomStatus", "Maintenance"
    SQLText "AND Active = TRUE", False, True
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("SUM5").Value > 0 Then
            intSum5 = CInt(rst("SUM5").Value)
        Else
            intSum5 = 0
        End If
    Else
        intSum5 = 0
    End If
    lblSummary5.Caption = "Maintenance : " & intSum5
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

Public Sub ShowSummary4()
    Const mstrMethod As String = "ShowSummary4"
    Dim rst As ADODB.Recordset
    Dim intSum4 As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "COUNT(ID) AS SUM4", False
    SQL_FROM "Room"
    SQL_WHERE_Text "RoomStatus", "Housekeeping"
    SQLText "AND Active = TRUE", False
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("SUM4").Value > 0 Then
            intSum4 = CInt(rst("SUM4").Value)
        Else
            intSum4 = 0
        End If
    Else
        intSum4 = 0
    End If
    lblSummary4.Caption = "Housekeeping : " & intSum4
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

Public Sub ShowSummary3()
    Const mstrMethod As String = "ShowSummary3"
    Dim rst As ADODB.Recordset
    Dim intSum3 As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "COUNT(ID) AS SUM3", False
    SQL_FROM "Room"
    SQL_WHERE_Text "RoomStatus", "Occupied"
    SQLText "AND Active = TRUE", False
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("SUM3").Value > 0 Then
            intSum3 = CInt(rst("SUM3").Value)
        Else
            intSum3 = 0
        End If
    Else
        intSum3 = 0
    End If
    lblSummary3.Caption = "Occupied : " & intSum3
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

Public Sub ShowSummary2()
    Const mstrMethod As String = "ShowSummary2"
    Dim rst As ADODB.Recordset
    Dim intSum2 As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "COUNT(ID) AS SUM2", False
    SQL_FROM "Room"
    SQL_WHERE_Text "RoomStatus", "Booked"
    SQLText "AND Active = TRUE", False
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("SUM2").Value > 0 Then
            intSum2 = CInt(rst("SUM2").Value)
        Else
            intSum2 = 0
        End If
    Else
        intSum2 = 0
    End If
    lblSummary2.Caption = "Booked : " & intSum2
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

Public Sub ShowSummary1()
    Const mstrMethod As String = "ShowSummary1"
    Dim rst As ADODB.Recordset
    Dim intSum1 As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "COUNT(ID) AS SUM1", False
    SQL_FROM "Room"
    SQL_WHERE_Text "RoomStatus", "Open"
    SQLText "AND Active = TRUE", False
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("SUM1").Value > 0 Then
            intSum1 = CInt(rst("SUM1").Value)
        Else
            intSum1 = 0
        End If
    Else
        intSum1 = 0
    End If
    lblSummary1.Caption = "Open : " & intSum1
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

Public Sub SetButtonProperties()
    Const mstrMethod As String = "SetButtonProperties"
    Dim rst As ADODB.Recordset
    Dim intID As Integer
    Dim strRoomType As String
    Dim strStatus As String
    Dim i As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "R.ID"
    SQLText "R.RoomShortName"
    SQLText "R.RoomType"
    SQLText "R.Maintenance"
    SQLText "R.RoomStatus"
    SQLText "R.Active", False
    SQL_FROM "Room", "R"
    SQL_LEFT_JOIN "Booking", "B"
    SQL_ON "R", "BookingID", "B", "ID"
    'SQLText "WHERE B.Status <> 'Closed'", False
    
    ' Reset blink
    For i = 1 To 61
        blnBlink(i) = False
    Next
    OpenDB
    Set rst = OpenRS(gstrSQL)
    While Not rst.EOF
        If rst("ID").Value > 0 And rst("ID").Value < cmdUnit.UBound + 1 Then
            intID = CInt(rst("ID").Value)
            If rst("RoomType").Value <> "" Then
                strRoomType = Trim(rst("RoomType").Value)
            Else
                strRoomType = ""
            End If
            'If intID > 22 Then 'If strRoomType <> "DORM" Then
                cmdUnit(intID).Caption = Trim(rst("RoomShortName").Value) & vbCrLf & vbCrLf & strRoomType
            'Else
            '    cmdUnit(intID).Caption = Trim(rst("RoomShortName").Value)
            'End If
'            If rst("Maintenance").Value = True Then
'                strStatus = "Maintenance"
''                cmdUnit(intID).BackColor = COL_BLUE
''                blnBlink(intID) = False
'            Else
                If rst("RoomStatus").Value <> "" Then
                    strStatus = Trim(rst("RoomStatus").Value)
                Else
                    strStatus = "Open"
                End If
'            End If
            Select Case strStatus
                Case "Maintenance"
                    cmdUnit(intID).BackColor = COL_BLUE
                    blnBlink(intID) = False
                Case "Housekeeping"
                    cmdUnit(intID).BackColor = COL_PURPLE
                    blnBlink(intID) = False
                Case "Booked"
                    cmdUnit(intID).BackColor = COL_YELLOW
                    blnBlink(intID) = AlertBooking(intID, strStatus)
                Case "Occupied"
                    cmdUnit(intID).BackColor = COL_RED
                    blnBlink(intID) = AlertBooking(intID, strStatus)
                'Case ""
                '    cmdUnit(intID).BackColor = COL_GREEN
                Case Else ' No Colour
                    cmdUnit(intID).BackColor = COL_GREEN '&HC0C0C0
                    blnBlink(intID) = False
            End Select
            If rst("Active").Value = True Then
                cmdUnit(intID).Visible = True
            Else
                cmdUnit(intID).Visible = False
            End If
        End If
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

Private Function GetRoomStatus(intRoomID As Integer) As String
    Const mstrMethod As String = "GetRoomStatus"
    Dim rst As ADODB.Recordset
    Dim mstrStatus As String
    'Dim blnHousekeeping As Boolean
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "RoomStatus", False
    SQL_FROM "Room"
    SQL_WHERE_Integer "ID", intRoomID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("RoomStatus").Value <> "" Then
            mstrStatus = Trim(rst("RoomStatus").Value)
        Else
            mstrStatus = ""
        End If
    Else
        mstrStatus = ""
    End If
    GetRoomStatus = mstrStatus
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

Private Function GetBookingID(intRoomID As Integer) As Long
    Const mstrMethod As String = "GetBookingID"
    Dim rst As ADODB.Recordset
    Dim mlngBookingID As Integer
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "BookingID", False
    SQL_FROM "Room"
    SQL_WHERE_Integer "ID", intRoomID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst("BookingID").Value <> "" Then ' Not Null
            mlngBookingID = ConvLng(rst("BookingID").Value)
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

Private Function RoomSetup(intRoomID As Integer) As Boolean
    Const mstrMethod As String = "RoomSetup"
    Dim rst As ADODB.Recordset
    Dim blnSetup As Boolean
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "ID", False
    SQL_FROM "Room"
    SQL_WHERE_Integer "ID", intRoomID
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        blnSetup = True
    Else
        blnSetup = False
    End If
    RoomSetup = blnSetup
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

Private Function AlertBooking(intRoomID As Integer, strStatus As String) As Boolean
    Const mstrMethod As String = "Alert Booking"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    SQL_SELECT
    If strStatus = "Booked" Then
        SQLText "B.GuestCheckIN AS CheckDate", False
    Else ' Occupied
        SQLText "B.GuestCheckOUT AS CheckDate", False
    End If
    SQL_FROM "Room", "R"
    SQL_LEFT_JOIN "Booking", "B"
    SQL_ON "R", "BookingID", "B", "ID"
    If strStatus = "Booked" Then
        SQL_WHERE_Text "R.RoomStatus", "Booked"
    Else
        SQL_WHERE_Text "R.RoomStatus", "Occupied"
    End If
    SQLText "AND R.ID = " & intRoomID, False, True
    SQLText "AND B.Active = TRUE", False, True
    OpenDB
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If Now > rst("CheckDate").Value Then
            AlertBooking = True
            If strStatus = "Booked" Then
                strBackColor(intRoomID) = COL_YELLOW
            Else
                strBackColor(intRoomID) = COL_RED
            End If
        Else
            AlertBooking = False
        End If
    Else
        AlertBooking = False
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Private Sub BlinkButton()
    tmrBlink.Enabled = Not tmrBlink.Enabled
    UpdateBlinkSetting tmrBlink.Enabled
    With tbrMenu.Buttons("BLINK")
        If .Caption = "Blink (F7)" Then
            .Caption = "Unblink (F7)"
            .ToolTipText = "Unblink Button (F7)"
            .Image = 8
        Else
            .Caption = "Blink (F7)"
            .ToolTipText = "Blink Button (F7)"
            .Image = 7
        End If
    End With
    SetButtonProperties
End Sub

Public Function NeedBlink() As Boolean
    Dim i As Integer
    For i = 0 To 61
        If blnBlink(i) = True Then
            NeedBlink = True
            Exit Function
        End If
    Next
    NeedBlink = False
End Function

Private Sub LoadBlinkSetting()
    Const mstrMethod As String = "LoadBlinkSetting"
    Dim rst As ADODB.Recordset
    Dim blnBlinkSetting As Boolean
    Dim blnNeedBlink As Boolean
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "DashboardBlink", False
    SQL_FROM "UserData"
    SQL_WHERE_Text "UserID", gstrUserID
    OpenDB
    With tbrMenu.Buttons("BLINK")
    Set rst = OpenRS(gstrSQL)
    If Not rst.EOF Then
        If rst!DashboardBlink = True Then
            blnBlinkSetting = True
            .Caption = "Unblink (F7)"
            .ToolTipText = "Unblink Button (F7)"
            .Image = 8
        Else
            blnBlinkSetting = False
            .Caption = "Blink (F7)"
            .ToolTipText = "Blink Button (F7)"
            .Image = 7
        End If
    End If
    CloseRS rst
    End With
    CloseDB
    blnNeedBlink = NeedBlink
    If blnNeedBlink And blnBlinkSetting Then
        tmrBlink.Enabled = True
    End If
    Exit Sub
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub UpdateBlinkSetting(blnValue As Boolean)
    Const mstrMethod As String = "UpdateBlinkSetting"
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    OpenDB
    SQL_UPDATE "UserData"
    SQL_SET_Boolean "DashboardBlink", blnValue, False
    SQL_WHERE_Text "UserID", gstrUserID
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub
