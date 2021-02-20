VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preview"
   ClientHeight    =   9645
   ClientLeft      =   4845
   ClientTop       =   2955
   ClientWidth     =   15375
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tbrButton 
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   1429
      ButtonWidth     =   1905
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Close (F4)"
            Key             =   "CLOSE"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Refresh (F5)"
            Key             =   "REFRESH"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Settings (F6)"
            Key             =   "SETUP"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Export (F7)"
            Key             =   "EXPORT"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Print (F8)"
            Key             =   "PRINT"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmPrint.frx":08CA
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   9840
         Top             =   120
      End
      Begin VB.Frame fraLogin 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10320
         TabIndex        =   2
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
            TabIndex        =   4
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
            TabIndex        =   3
            Top             =   390
            Width           =   1620
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6000
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrint.frx":0BE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrint.frx":14BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrint.frx":1DB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrint.frx":268A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrint.frx":2F64
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   10695
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   1815
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   10695
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   0   'False
         EnableZoomControl=   -1  'True
         EnableCloseButton=   0   'False
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   0   'False
         EnableDrillDown =   0   'False
         EnableAnimationControl=   0   'False
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   0   'False
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmPrint.frx":3E26
      Top             =   120
      Width           =   4050
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
      TabIndex        =   6
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
      TabIndex        =   5
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
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 27/12/2014
' Descriptions : 1) Use timer to count after an interval then auto log out
Option Explicit
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const mstrModule As String = "Print"
Dim CrReport As CRAXDRT.Report
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
'Dim ESC_Pressed As Boolean
Dim F4_Pressed As Boolean
Dim F5_Pressed As Boolean
Dim F6_Pressed As Boolean
Dim F7_Pressed As Boolean
Dim F8_Pressed As Boolean
Dim intTick As Integer

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyP Then
'        MsgBox "P pressed"
'    End If
'End Sub

Private Sub Timer1_Timer()
    'ESC_Pressed = (GetKeyState(vbKeyEscape) < 0)
    F4_Pressed = (GetKeyState(vbKeyF4) < 0)
    F5_Pressed = (GetKeyState(vbKeyF5) < 0)
    F6_Pressed = (GetKeyState(vbKeyF6) < 0)
    F7_Pressed = (GetKeyState(vbKeyF7) < 0)
    F8_Pressed = (GetKeyState(vbKeyF8) < 0)
    'Cls
    'Debug.Print "{Esc} pressed = " & ESC_Pressed & "   {F4} pressed = " & F4_Pressed
    If F4_Pressed Then
        Unload Me
    ElseIf F5_Pressed And tbrButton.Buttons("REFRESH").Enabled Then
        CRViewer1.Refresh
    ElseIf F6_Pressed And tbrButton.Buttons("SETUP").Enabled Then
        CrReport.PrinterSetup (0)
    ElseIf F7_Pressed And tbrButton.Buttons("EXPORT").Enabled Then
        CrReport.Export
        CRViewer1.Refresh
    ElseIf F8_Pressed And tbrButton.Buttons("PRINT").Enabled Then
        CRViewer1.PrintReport
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
    Const mstrMethod As String = "Loading Print"
    Screen.MousePointer = vbHourglass
On Error GoTo CheckErr
    lblSystemDateTime.Caption = FormatDateDayAndTime(Now)
    lblUserID.Caption = "User ID : " & gstrUserID
    Me.Caption = gstrReportTitle
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 680
    If UserAccessModule(MOD_REPORT_EXPORT) = True Then
        tbrButton.Buttons("EXPORT").Enabled = True
    Else
        tbrButton.Buttons("EXPORT").Enabled = False
    End If
    Set CrReport = CrApplication.OpenReport(App.Path & "\Report\" & gstrReportFileName)  ', 1)
    OpenRDB
    rs.Open gstrSQL, Con, adOpenStatic, adLockPessimistic
    CrReport.Database.SetDataSource rs, 3, 1
    CRViewer1.ReportSource = CrReport
    CrReport.ReportTitle = gstrReportTitle
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    Screen.MousePointer = vbDefault
    Exit Sub
CheckErr:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
    Const mstrMethod As String = "Print Form_Resize"
On Error GoTo CheckErr
'    fraPanel.Top = 2040
'    fraPanel.Left = 0
'    If Me.Height - 2500 > 0 Then
'        fraPanel.Height = Me.Height - 2500
'        CRViewer1.Height = fraPanel.Height '- 2500
'    End If
'    If Me.Width - 100 > 0 Then
'        fraPanel.Width = Me.Width '- 100
'        CRViewer1.Width = fraPanel.Width - 100
'    End If
    ' If Form Minimize?
    'If Not Me.WindowState = vbMinimized Then
'    CRViewer1.Top = 2040 '810
'    CRViewer1.Left = 0
'    If Me.Height - 2500 > 0 Then
'        CRViewer1.Height = Me.Height - 2500 ' ScaleHeight - 810
'    End If
'    If Me.Width - 100 > 0 Then
'        CRViewer1.Width = Me.Width - 100 ' ScaleWidth
'    End If
    'End If
    
    Frame1.Top = 2040
    Frame1.Left = 0
    'CRViewer1.Top = 2040 '810
    'CRViewer1.Left = 0
    If Me.Height - 2500 > 0 Then
        Frame1.Height = Me.Height - 2500
        CRViewer1.Height = Me.Height - 2500 ' ScaleHeight - 810
    End If
    If Me.Width - 100 > 0 Then
        Frame1.Width = Me.Width - 100
        CRViewer1.Width = Me.Width - 100 ' ScaleWidth
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const mstrMethod As String = "Print Form_Unload"
On Error GoTo CheckErr
    CloseRS rs
    CloseRDB
    Set CrReport = Nothing
    'Set CrApplication = Nothing
    'frmDashboard.Show
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        Unload Me
'    ElseIf KeyCode = vbKeyF5 And tbrButton.Buttons("REFRESH").Enabled Then
'        CRViewer1.Refresh
'    ElseIf KeyCode = vbKeyF6 And tbrButton.Buttons("SETUP").Enabled Then
'        CrReport.PrinterSetup (0)
'    ElseIf KeyCode = vbKeyF4 And tbrButton.Buttons("EXPORT").Enabled Then
'        CrReport.Export
'        CRViewer1.Refresh
'    ElseIf KeyCode = vbKeyP And (Shift And vbCtrlMask) And tbrButton.Buttons("PRINT").Enabled Then
'        CRViewer1.PrintReport
'    End If
'End Sub

Private Sub tbrButton_ButtonClick(ByVal Button As MSComctlLib.Button)
    Const mstrMethod As String = "Print tbrButton_ButtonClick"
On Error GoTo CheckErr
    Select Case Button.Key
        Case "CLOSE"
            Unload Me
        Case "PRINT"
            CRViewer1.PrintReport
        Case "EXPORT"
            CrReport.Export
            CRViewer1.Refresh
        Case "REFRESH"
            CRViewer1.Refresh
        Case "SETUP"
            CrReport.PrinterSetup (0)
    End Select
    Exit Sub
CheckErr:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub OpenRDB()
    Const mstrMethod As String = "Print OpenRDB"
On Error GoTo CheckErr
    Con.Provider = "Microsoft.Jet.OLEDB.4.0"
    Con.ConnectionString = "Data Source=" & gstrDatabasePath
    Con.Properties("Jet OLEDB:Database Password") = gstrPassword
    Con.Open
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Public Sub CloseRDB()
    Const mstrMethod As String = "Print CloseRDB"
On Error GoTo CheckErr
    If Con Is Nothing Then
    Else
        If Con.State = adStateOpen Then
            Con.Close
        End If
        Set Con = Nothing
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub
