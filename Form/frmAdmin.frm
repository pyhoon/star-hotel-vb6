VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00303030&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admin Access"
   ClientHeight    =   2790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8175
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUserID 
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF7929&
      Height          =   435
      Left            =   3480
      MaxLength       =   10
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      ToolTipText     =   "Enter your User ID"
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0076E600&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3480
      MaxLength       =   10
      MousePointer    =   1  'Arrow
      PasswordChar    =   "n"
      TabIndex        =   3
      ToolTipText     =   "Enter Password"
      Top             =   960
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
      MouseIcon       =   "frmAdmin.frx":1D42
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1800
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
      Left            =   4080
      MouseIcon       =   "frmAdmin.frx":2A0C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1800
      Width           =   1620
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
      Picture         =   "frmAdmin.frx":36D6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password :"
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
      TabIndex        =   5
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User ID :"
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
      TabIndex        =   4
      Top             =   360
      Width           =   1560
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Added On : 04/01/2014
' Descriptions : 1) Temporary Admin login to allow Void
Option Explicit
Private Const mstrModule As String = "Admin Access"

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Const mstrMethod As String = "cmdOK_Click"
    Dim strUser As String
    Dim strPassword As String
    Dim intAttempt As Integer
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    If Trim(txtUserID.Text) = "" Then
        MsgBox "Please enter your User ID!", vbExclamation, mstrModule
        txtUserID.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        MsgBox "Please enter your Password!", vbExclamation, mstrModule
        txtPassword.SetFocus
        Exit Sub
    End If
    SQL_SELECT
    SQLText "UserID"
    SQLText "UserGroup"
    SQLText "UserPassword"
    SQLText "Salt"
    SQLText "Active"
    SQLText "LoginAttempts", False
    SQL_FROM "UserData"
    SQL_WHERE_Text "UserID", CheckInput(txtUserID.Text)
    OpenDB
    Set rst = OpenSQL(gstrSQL)
    If rst.RecordCount = 0 Then
        CloseRS rst
        CloseDB
        Exit Sub
    End If
    'check for correct password
    If Not rst.EOF Then
        '==============================================================================
        strUser = rst!UserID
        If rst!Active = False Then
            MsgBox "Your User ID has been frozen." & vbCrLf & _
                "Please contact Super User.", vbExclamation, mstrModule
            CloseRS rst
            CloseDB
            Exit Sub
        End If
        '==============================================================================
        If ConvInt(rst!LoginAttempts) > 2 Then
            If rst!Active = True Then
                UpdateField "UserData", "Active", "BOOLEAN", "TRUE", " UserID = '" & strUser & "'"
            End If
            MsgBox "Your User ID has been frozen." & vbCrLf & _
                    "Please contact Super User.", vbExclamation, mstrModule
            CloseRS rst
            CloseDB
            Exit Sub
        End If
        '==============================================================================
        'Encrypt password so we can check it against the encypted password in the database
        'Read in the salt
        strPassword = txtPassword.Text
        strPassword = strPassword & rst!Salt
    
        'Encrypt the entered password
        strPassword = GoldFishEncode(strPassword)
            
        If strPassword = rst!UserPassword Then
            intAttempt = 0
            UpdateField "UserData", "Loginattempts", "NUMBER", "0", " UserID = '" & strUser & "'"
            CloseRS rst
            CloseDB

'            ' Check Module Access
'            If UserAccessModule(MOD_BOOKING_VOID, strUser) Then
'                frmBooking.VoidBooking CLng(lblBookingID.Caption)
'                Unload Me
'            Else
'                MsgBox "Your have no access to this function!", vbExclamation, mstrModule
'                Unload Me
'            End If
            Exit Sub
        Else
            If rst!UserGroup > 1 Then
                intAttempt = intAttempt + 1
                UpdateField "UserData", "Loginattempts", "NUMBER", "Loginattempts + 1", " UserID = '" & strUser & "'"
                CloseRS rst
                CloseDB
                If intAttempt > 2 Then
                    MsgBox "Too many attempts.", vbExclamation, mstrModule
                    Unload Me
                End If
            End If
        End If
    End If
    
    '==============================================================================
    MsgBox "Invalid Password, please try again!", vbExclamation, mstrModule
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
CheckErr:
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub
