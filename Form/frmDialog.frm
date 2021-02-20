VERSION 5.00
Begin VB.Form frmDialog 
   BackColor       =   &H00303030&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log out"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Cancel (Esc)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MouseIcon       =   "frmDialog.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2280
      Width           =   1620
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK (Enter)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      MouseIcon       =   "frmDialog.frx":110C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Timer tmrClock 
      Interval        =   1000
      Left            =   5400
      Top             =   2400
   End
   Begin VB.Label lblInstruction 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click OK to log out now or Cancel to ignore."
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
      Height          =   240
      Left            =   1230
      TabIndex        =   4
      Top             =   1680
      Width           =   3825
   End
   Begin VB.Label lblSecondsLeft 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 second(s)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   2205
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alert: System will be log out automatically after"
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
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   5325
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Added On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
Option Explicit
Dim intSecondsLeft As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
    CloseAllPrintForms
    Unload frmReportMaintain
    Unload frmReport
    Unload frmRoomTypeMaintain
    Unload frmRoomMaintain
    Unload frmUserMaintain
    Unload frmModuleAccess
    Unload frmBooking
    Unload frmDashboard
'    CloseAllForms
End Sub

Private Sub Form_Load()
    intSecondsLeft = 10
    lblSecondsLeft.Caption = "10 second(s)"
End Sub

Private Sub tmrClock_Timer()
    intSecondsLeft = intSecondsLeft - 1
    lblSecondsLeft.Caption = intSecondsLeft & " second(s)"
    If intSecondsLeft < 1 Then
        Unload Me
        CloseAllPrintForms
        Unload frmReportMaintain
        Unload frmReport
        Unload frmRoomTypeMaintain
        Unload frmRoomMaintain
        Unload frmUserMaintain
        Unload frmModuleAccess
        Unload frmBooking
        Unload frmDashboard
    '    CloseAllForms
    End If
End Sub

Private Sub CloseAllPrintForms()
    Dim frm As Form
    For Each frm In Forms
        If frm.Name = "frmPrint" Then
            Unload frm
            Set frm = Nothing
        End If
    Next
End Sub

Private Sub CloseAllForms()
    Dim frm As Form
    Unload frmDialog
    Set frmDialog = Nothing
    For Each frm In Forms
        If Not (frm.Name = "frmDialog" Or frm.Name = "frmDashboard") Then
            Unload frm
            Set frm = Nothing
        End If
    Next
    Unload frmDashboard
    Set frmDashboard = Nothing
End Sub
