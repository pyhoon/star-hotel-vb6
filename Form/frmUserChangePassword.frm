VERSION 5.00
Begin VB.Form frmUserChangePassword 
   BackColor       =   &H00303030&
   BorderStyle     =   0  'None
   Caption         =   "Change Password"
   ClientHeight    =   5790
   ClientLeft      =   2790
   ClientTop       =   3105
   ClientWidth     =   11070
   Icon            =   "frmUserChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420.922
   ScaleMode       =   0  'User
   ScaleWidth      =   10394.13
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   240
      ScaleHeight     =   5055
      ScaleWidth      =   10575
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   10575
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   6720
         MouseIcon       =   "frmUserChangePassword.frx":1D42
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   3600
         Width           =   1620
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   8640
         MouseIcon       =   "frmUserChangePassword.frx":2A0C
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   3600
         Width           =   1620
      End
      Begin VB.TextBox txtPasswordNew 
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
         Left            =   6000
         MaxLength       =   10
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "n"
         TabIndex        =   1
         ToolTipText     =   "Enter New Password"
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txtPasswordOld 
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
         ForeColor       =   &H00FF7929&
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   6000
         MaxLength       =   10
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "n"
         TabIndex        =   0
         ToolTipText     =   "Enter your Old Password"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtPasswordConfirm 
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
         ForeColor       =   &H004417FF&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   6000
         MaxLength       =   10
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "n"
         TabIndex        =   2
         ToolTipText     =   "Re-enter New Password"
         Top             =   2880
         Width           =   4215
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Old Password:"
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
         Left            =   3000
         TabIndex        =   8
         Top             =   1680
         Width           =   2760
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&New Password:"
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
         Left            =   3000
         TabIndex        =   9
         Top             =   2280
         Width           =   2760
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   1440
         Picture         =   "frmUserChangePassword.frx":36D6
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   765
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   2430
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   360
         TabIndex        =   6
         Top             =   0
         Width           =   3000
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00505050&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   5880
         Top             =   1665
         Width           =   4455
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00505050&
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   5880
         Top             =   2265
         Width           =   4455
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00505050&
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   5880
         Top             =   2865
         Width           =   4455
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Confirm Password:"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   10
         Top             =   2880
         Width           =   2760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   -120
         X2              =   10680
         Y1              =   4560
         Y2              =   4560
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
         Left            =   0
         TabIndex        =   11
         Top             =   4680
         Width           =   10500
      End
   End
End
Attribute VB_Name = "frmUserChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 19/05/2018
Option Explicit
Private Const mstrModule As String = "Change Password"
Dim strPassword As String
Dim intAttempt As Integer
Dim blnChange As Boolean

Private Sub Form_Load()
    'Me.BackColor = &HFFC000
    'fraLogin.BackColor = &HFFC000
    lblCompanyProduct.Caption = COMPANY_PRODUCT_NAME
    lblProductName.Caption = App.ProductName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If UserAccessModule(MOD_DASHBOARD) Then
        frmDashboard.Show
        Unload Me
    Else
        MsgBox "Your access has been disabled!", vbExclamation, mstrModule
        'frmBooking.Show
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    Const mstrMethod As String = "Form_Resize"
On Error GoTo CheckErr
    fraLogin.Left = (Me.Width - fraLogin.Width) \ 2
    fraLogin.Top = (Me.Height - fraLogin.Height) \ 2 - fraLogin.Height
'    Line1.X1 = 150
'    Line1.X2 = Me.Width - 1500
'    Line1.Y1 = Me.Height - 5000
'    Line1.Y2 = Line1.Y1
'    lblCopyright.Left = 300
'    lblCopyright.Width = Me.Width - 2000
'    lblCopyright.Top = Me.Height - 4900
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub cmdOK_Click()
    Const mstrMethod As String = "cmdOK_Click"
    Dim strPass As String
    Dim strNew As String
On Error GoTo CheckErr
    If Trim(txtPasswordOld.Text) = "" Then
        MsgBox "Please enter your Old Password!", vbExclamation, mstrModule
        txtPasswordOld.SetFocus
        Exit Sub
    End If
    If Trim(txtPasswordNew.Text) = "" Then
        MsgBox "Please enter your New Password!", vbExclamation, mstrModule
        txtPasswordNew.SetFocus
        Exit Sub
    End If
    strPass = CheckInput(txtPasswordOld.Text)
    If CheckPassword(strPass) = False Then
        MsgBox "Old Password is incorrect!", vbExclamation, mstrModule
        Exit Sub
    End If
'    If txtPasswordNew.Text = "" Then
'        MsgBox "Password cannot empty!", vbExclamation, mstrModule
'        txtPasswordNew.SetFocus
'        Exit Sub
'    End If
    If Len(txtPasswordNew.Text) < 4 Then
        MsgBox "Password must at least 4 characters!", vbExclamation, mstrModule
        txtPasswordNew.SetFocus
        Exit Sub
    End If
    If txtPasswordNew.Text = txtPasswordConfirm.Text Then
        strNew = CheckInput(txtPasswordNew.Text)
        UpdatePassword strNew
    Else
        MsgBox "Please confirm Password!", vbExclamation, mstrModule
        txtPasswordConfirm.SetFocus
        Exit Sub
    End If
    If UserAccessModule(MOD_DASHBOARD) Then
        frmDashboard.Show
        Unload Me
    Else
        MsgBox "Your access has been disabled!", vbExclamation, mstrModule
        'frmBooking.Show
        Unload Me
    End If
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub cmdCancel_Click()
    If gblnUserChangePassword Then
        Unload Me
        frmDashboard.Show
        gblnUserChangePassword = False
    Else
        End
    End If
End Sub

Private Function CheckPassword(strPassword As String) As Boolean
    Const mstrMethod As String = "CheckPassword"
    Dim rst As ADODB.Recordset
    Dim strCheck As String
On Error GoTo CheckErr
    SQL_SELECT
    SQLText "UserPassword"
    SQLText "Salt", False
    SQL_FROM "UserData"
    SQL_WHERE_Text "UserID", gstrUserID
    OpenDB
    Set rst = OpenSQL(gstrSQL)
    If Not rst.EOF Then
        strCheck = strPassword & rst!Salt
        strCheck = GoldFishEncode(strCheck)
        If rst!UserPassword = strCheck Then
            CheckPassword = True
        Else
            CheckPassword = False
        End If
    Else
        CheckPassword = False
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Private Sub UpdatePassword(strPassword As String)
    Const mstrMethod As String = "UpdatePassword"
    Dim mstrSalt As String
On Error GoTo CheckErr
    mstrSalt = GenSalt(4)
    SQL_UPDATE "UserData"
    SQL_SET_Text "UserPassword", Encrypt(strPassword, mstrSalt)
    SQL_SET_Text "Salt", mstrSalt
    SQL_SET_Boolean "ChangePassword", False, False
    SQL_WHERE_Text "UserID", gstrUserID
    OpenDB
    QuerySQL gstrSQL
    CloseDB
    MsgBox "Password is updated!", vbInformation, mstrModule
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub
