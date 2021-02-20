VERSION 5.00
Begin VB.Form frmDatabase 
   BackColor       =   &H00303030&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database File"
   ClientHeight    =   3480
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8175
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDemo 
      BackColor       =   &H00303030&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Use DemoData.mdb"
      Top             =   1680
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.TextBox txtFilePath 
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Enter File Path"
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Enter File Name"
      Top             =   1080
      Width           =   4215
   End
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
      Left            =   6120
      MouseIcon       =   "frmDatabase.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2520
      Width           =   1620
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save (Enter)"
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
      Left            =   4080
      MouseIcon       =   "frmDatabase.frx":1254
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2520
      Width           =   1620
   End
   Begin VB.Label lblBack 
      BackColor       =   &H00303030&
      Caption         =   "Use Demo database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1720
      Width           =   4095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "File &Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label lblBookingID 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00505050&
      BackStyle       =   1  'Opaque
      Height          =   525
      Left            =   3360
      Top             =   945
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00505050&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3360
      Top             =   345
      Width           =   4455
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   360
      Picture         =   "frmDatabase.frx":1F1E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "File &Path :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   1560
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Added On : 13/07/2019
' Descriptions : 1) Custom database location
Option Explicit
Private Const mstrModule As String = "Database File"

Private Sub Form_Load()
    If gstrDatabasePath = App.Path & "\Data\DemoData.mdb" Then
        chkDemo.Value = vbChecked
        txtFilePath.Enabled = False
        txtFileName.Enabled = False
    Else
        chkDemo.Value = vbUnchecked
        txtFilePath.Enabled = True
        txtFileName.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    frmUserLogin.Show
End Sub

Private Sub cmdOK_Click()
    Const mstrMethod As String = "cmdOK_Click"
    Dim strError As String
    Dim strPath As String
    Dim strFile As String
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    If chkDemo.Value = vbChecked Then
        'gstrDatabasePath = App.Path & "\Data\DemoData.mdb"
        If FileExists(App.Path & "\Data\DemoData.mdb") = False Then
            MsgBox "DemoData.mdb is not found!", vbExclamation, mstrModule
            Exit Sub
        End If
        strPath = App.Path & "\Data\"
        strFile = "DemoData.mdb"
    Else
        If Trim(txtFilePath.Text) = "" Then
            MsgBox "Please enter File Path!", vbExclamation, mstrModule
            txtFilePath.SetFocus
            Exit Sub
        End If
        If Trim(txtFileName.Text) = "" Then
            MsgBox "Please enter File Name!", vbExclamation, mstrModule
            txtFileName.SetFocus
            Exit Sub
        End If
        strPath = Trim(txtFilePath.Text)
        strFile = Trim(txtFileName.Text)
    End If

    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    Dim strOutput As String
    strOutput = strPath & vbCrLf & strFile
    WriteTextFile App.Path & "\Config.txt", strOutput
    
    gstrDatabasePath = strPath & strFile
    If FileExists(gstrDatabasePath) Then
        Unload Me
        frmSplash.Show
        Exit Sub
    End If
    
    If CreateData = False Then
        strError = "Error during generating database file."
        MsgBox strError, vbExclamation, mstrModule
        Exit Sub
    End If
    
    If CreateDB = False Then
        strError = "Error during creating database file."
        MsgBox strError, vbExclamation, mstrModule
        Exit Sub
    End If
    
    If CreateSampleData = False Then
        strError = "Error during inserting sample data."
        MsgBox strError, vbExclamation, mstrModule
        Exit Sub
    End If
    
    Unload Me
    frmSplash.Show
    Exit Sub
CheckErr:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
End Sub

Private Sub chkDemo_Click()
    If chkDemo.Value = vbChecked Then
        txtFilePath.Enabled = False
        txtFileName.Enabled = False
    Else
        txtFilePath.Enabled = True
        txtFileName.Enabled = True
    End If
End Sub
