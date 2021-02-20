VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModuleAccess 
   BackColor       =   &H00505050&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access Control"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15375
   Icon            =   "frmModuleAccess.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7485
      Left            =   240
      ScaleHeight     =   7485
      ScaleWidth      =   7575
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7575
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   1
         Top             =   525
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   11
         Left            =   4200
         TabIndex        =   11
         Top             =   4125
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   11
         Left            =   6240
         TabIndex        =   29
         Top             =   4125
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   18
         Left            =   6240
         TabIndex        =   36
         Top             =   7020
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   18
         Left            =   4200
         TabIndex        =   18
         Top             =   7020
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   17
         Left            =   6240
         TabIndex        =   35
         Top             =   6660
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   16
         Left            =   6240
         TabIndex        =   34
         Top             =   6300
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   15
         Left            =   6240
         TabIndex        =   33
         Top             =   5940
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   14
         Left            =   6240
         TabIndex        =   32
         Top             =   5565
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   13
         Left            =   6240
         TabIndex        =   31
         Top             =   5220
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   12
         Left            =   6240
         TabIndex        =   30
         Top             =   4860
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   10
         Left            =   6240
         TabIndex        =   28
         Top             =   3765
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   9
         Left            =   6240
         TabIndex        =   27
         Top             =   3420
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   8
         Left            =   6240
         TabIndex        =   26
         Top             =   3045
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   7
         Left            =   6240
         TabIndex        =   25
         Top             =   2685
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   6
         Left            =   6240
         TabIndex        =   24
         Top             =   2325
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   5
         Left            =   6240
         TabIndex        =   23
         Top             =   1965
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   22
         Top             =   1605
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   21
         Top             =   1245
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   20
         Top             =   885
         Width           =   615
      End
      Begin VB.CheckBox chkGroup4 
         BackColor       =   &H00303030&
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   19
         Top             =   525
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   17
         Left            =   4200
         TabIndex        =   17
         Top             =   6660
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   16
         Left            =   4200
         TabIndex        =   16
         Top             =   6300
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   15
         Left            =   4200
         TabIndex        =   15
         Top             =   5940
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   14
         Left            =   4200
         TabIndex        =   14
         Top             =   5565
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   13
         Left            =   4200
         TabIndex        =   13
         Top             =   5220
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   12
         Left            =   4200
         TabIndex        =   12
         Top             =   4860
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   10
         Top             =   3765
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   9
         Top             =   3420
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   8
         Top             =   3045
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   7
         Top             =   2685
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   6
         Top             =   2325
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   5
         Top             =   1965
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   4
         Top             =   1605
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   3
         Top             =   1245
         Width           =   615
      End
      Begin VB.CheckBox chkGroup1 
         BackColor       =   &H00303030&
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   2
         Top             =   885
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clerk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   65
         Top             =   4440
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3360
         TabIndex        =   66
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label lblGroup4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clerk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5400
         TabIndex        =   40
         Top             =   120
         Width           =   2025
      End
      Begin VB.Label lblGroup1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3360
         TabIndex        =   39
         Top             =   120
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   3360
         X2              =   3360
         Y1              =   120
         Y2              =   7320
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   5400
         X2              =   5400
         Y1              =   120
         Y2              =   7320
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   64
         Top             =   4080
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   62
         Top             =   6240
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   61
         Top             =   3360
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   18
         Left            =   120
         TabIndex        =   60
         Top             =   6960
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   17
         Left            =   120
         TabIndex        =   59
         Top             =   6600
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   58
         Top             =   5880
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   57
         Top             =   5520
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   56
         Top             =   5160
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   55
         Top             =   4800
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   54
         Top             =   3720
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   53
         Top             =   3000
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   52
         Top             =   2640
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   51
         Top             =   2280
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   50
         Top             =   1920
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Disabled Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   7305
      End
      Begin VB.Label lblModule 
         Appearance      =   0  'Flat
         BackColor       =   &H00303030&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   7305
      End
      Begin VB.Label lblReport 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   4440
         Width           =   7305
      End
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Height          =   810
      Left            =   0
      TabIndex        =   37
      Top             =   1200
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   1429
      ButtonWidth     =   1588
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close (Esc)"
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Close (Esc)"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmModuleAccess.frx":0CCA
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   42
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
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   390
            Width           =   1620
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   -2147483633
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModuleAccess.frx":0FE4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmModuleAccess.frx":18BE
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
      TabIndex        =   46
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
      TabIndex        =   45
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
Attribute VB_Name = "frmModuleAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
Option Explicit
Private Const mstrModule As String = "Module Access"
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
    lblBusinessName.Caption = COMPANY_PRODUCT_NAME
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    LoadModuleAccess
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDashboard.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "CLOSE"
            Unload Me
    End Select
End Sub

Private Sub LoadModuleAccess()
    Const mstrMethod As String = "LoadModuleAccess"
    Dim rst As ADODB.Recordset
    Dim i As Integer
On Error GoTo CheckErr
    SQL_SELECT_ALL "ModuleAccess"
    SQL_WHERE_Boolean "Active", True
    SQL_ORDER_BY "ModuleID"
    OpenDB
    Set rst = OpenRS(gstrSQL)
    While Not rst.EOF
        'i = i + 1
        i = rst!ModuleID
        If rst!ModuleID < 19 Then
            If rst!ModuleDesc1 <> "" Then
                lblModule(i).Caption = " " & rst!ModuleDesc1
            Else
                lblModule(i).Caption = ""
            End If
            If rst!Group1 Then
                chkGroup1(i).Value = vbChecked
            Else
                chkGroup1(i).Value = vbUnchecked
            End If
'            If rst!Group2 Then
'                chkGroup2(i).Value = vbChecked
'            Else
'                chkGroup2(i).Value = vbUnchecked
'            End If
'            If rst!Group3 Then
'                chkGroup3(i).Value = vbChecked
'            Else
'                chkGroup3(i).Value = vbUnchecked
'            End If
            If rst!Group4 Then
                chkGroup4(i).Value = vbChecked
            Else
                chkGroup4(i).Value = vbUnchecked
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

Private Sub UpdateModuleAccess(intModuleID As Integer, strGroup As String, blnValue As Boolean)
    Const mstrMethod As String = "UpdateModuleAccess"
    Dim lngRecordsAffected As Long
On Error GoTo CheckErr
    OpenDB
    SQL_UPDATE "ModuleAccess"
    SQL_SET_Boolean strGroup, blnValue, False
    SQL_WHERE_Integer "ModuleID", intModuleID
    QuerySQL gstrSQL, lngRecordsAffected
    CloseDB
    'MsgBox "Module Access updated! (" & lngRecordsAffected & ")", vbInformation, App.Title
    Exit Sub
CheckErr:
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub chkGroup1_Click(Index As Integer)
    If chkGroup1(Index).Value = vbChecked Then
        UpdateModuleAccess Index, "Group1", True
        'chkGroup1(Index).Caption = "Yes"
        'chkGroup1(Index).BackColor = &H80FF80
    Else
        UpdateModuleAccess Index, "Group1", False
        'chkGroup1(Index).Caption = "No"
        'chkGroup1(Index).BackColor = &H8080FF
    End If
End Sub

'Private Sub chkGroup2_Click(Index As Integer)
'    If chkGroup2(Index).Value = vbChecked Then
'        UpdateModuleAccess Index, "Group2", True
'        chkGroup2(Index).BackColor = &H80FF80
'    Else
'        UpdateModuleAccess Index, "Group2", False
'        chkGroup2(Index).BackColor = &H8080FF
'    End If
'End Sub

'Private Sub chkGroup3_Click(Index As Integer)
'    If chkGroup3(Index).Value = vbChecked Then
'        UpdateModuleAccess Index, "Group3", True
'        chkGroup3(Index).BackColor = &H80FF80
'    Else
'        UpdateModuleAccess Index, "Group3", False
'        chkGroup3(Index).BackColor = &H8080FF
'    End If
'End Sub

Private Sub chkGroup4_Click(Index As Integer)
    If chkGroup4(Index).Value = vbChecked Then
        UpdateModuleAccess Index, "Group4", True
        'chkGroup4(Index).Caption = "Yes"
        'chkGroup4(Index).BackColor = &H80FF80
    Else
        UpdateModuleAccess Index, "Group4", False
        'chkGroup4(Index).Caption = "No"
        'chkGroup4(Index).BackColor = &H8080FF
    End If
End Sub
