VERSION 5.00
Begin VB.Form frmUserLogin 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15375
   Icon            =   "frmUserLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15375
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox fraLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00303030&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5220
      Left            =   2880
      ScaleHeight     =   5220
      ScaleWidth      =   9375
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   9375
      Begin VB.CommandButton cmdOption 
         Caption         =   "Option"
         Height          =   615
         Left            =   7320
         MouseIcon       =   "frmUserLogin.frx":1D42
         MousePointer    =   99  'Custom
         TabIndex        =   13
         ToolTipText     =   "Change database file"
         Top             =   4440
         Width           =   1620
      End
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
         Left            =   5400
         MaskColor       =   &H00303030&
         MouseIcon       =   "frmUserLogin.frx":204C
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   3120
         UseMaskColor    =   -1  'True
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
         Left            =   7320
         MaskColor       =   &H00303030&
         MouseIcon       =   "frmUserLogin.frx":2D16
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   1620
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
         Left            =   4680
         MaxLength       =   10
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "n"
         TabIndex        =   1
         ToolTipText     =   "Enter Password"
         Top             =   2280
         Width           =   4215
      End
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
         Left            =   4680
         MaxLength       =   10
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         ToolTipText     =   "Enter your User ID"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label lblDemo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Width           =   405
      End
      Begin VB.Label lblVersionDB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Version"
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
         Left            =   120
         TabIndex        =   11
         Top             =   4440
         Width           =   1665
      End
      Begin VB.Label lblVersionApp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software Version"
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
         Left            =   120
         TabIndex        =   10
         Top             =   4080
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID :"
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
         Left            =   2880
         TabIndex        =   8
         Top             =   1680
         Width           =   1560
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   2280
         Width           =   1560
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   1560
         Picture         =   "frmUserLogin.frx":39E0
         Stretch         =   -1  'True
         Top             =   1560
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
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00505050&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   4560
         Top             =   1665
         Width           =   4455
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00505050&
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   4560
         Top             =   2265
         Width           =   4455
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
      TabIndex        =   5
      Top             =   900
      Width           =   5295
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   480
      Picture         =   "frmUserLogin.frx":8257
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
      TabIndex        =   4
      Top             =   240
      Width           =   15135
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
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Version : 1.2.22
'
' Modified On : 19/05/2018
' Modified On : 13/07/2019
Option Explicit
Private Const mstrModule As String = "User Login"
Dim strPassword As String
Dim intAttempt As Integer

Private Sub cmdOption_Click()
    Dim strPath As String
    Dim strFile As String
    Unload Me
    ReadTextFile "Config", 0, strPath
    ReadTextFile "Config", 1, strFile
    With frmDatabase
        .txtFilePath.Text = strPath
        .txtFileName.Text = strFile
        .Show vbModal
    End With
    Exit Sub
End Sub

Private Sub Form_Load()
    'Me.BackColor = RGB(184, 209, 255)
    'fraLogin.BackColor = RGB(184, 209, 255)
    lblBusinessName.Caption = COMPANY_PRODUCT_NAME
    lblCompanyProduct.Caption = COMPANY_PRODUCT_NAME
    lblProductName.Caption = App.ProductName
    lblVersionApp.Caption = "Software Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersionDB.Caption = "Database Version " & DB_Version
    'lblDemo.Caption = "This free demo is valid until " & DateAdd("D", 364, "01 Jan 2021") & " (Default User ID: admin && Password: admin)"
    lblDemo.Caption = "User ID: admin" & vbCrLf & "Password: admin"
End Sub

Private Sub cmdCancel_Click()
    'Unload frmBooking 'mdiMain
    'Unload frmSplash
    'Unload Me
    End
End Sub

Private Sub cmdOK_Click()
    Const mstrMethod As String = "cmdOK_Click"
    Dim rst As ADODB.Recordset
    Dim mblnActive As Integer
    Dim mintLoginAttempts As Integer
    Dim strSalt As String
On Error GoTo CheckErr
    ' Testing Purpose Only
'     If txtUserID.Text = "" And txtPassword.Text = "" Then
'         txtUserID.Text = "CLERK"
'         txtPassword.Text = "1"
'     End If
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
'    strPassword = CheckInput(txtPassword.Text)
    SQL_SELECT
    SQLText "UserGroup"
    SQLText "UserID"
    'SQLText "UserThemeID"
    SQLText "UserName"
    SQLText "UserPassword"
    SQLText "Salt"
    SQLText "Idle"
    SQLText "Active"
    SQLText "LoginAttempts"
    SQLText "ChangePassword", False
    SQL_FROM "UserData"
    SQL_WHERE_Text "UserID", CheckInput(txtUserID.Text)
    OpenDB
    Set rst = OpenSQL(gstrSQL)
    'If rst.RecordCount = 0 Then (-1)
    '    'mblnLoad = False
    '    CloseRS rst
    '    CloseDB
    '    Exit Sub
    'End If
    'check for correct password
    If rst.EOF Then
        CloseRS rst
        CloseDB
        MsgBox "User ID not found!", vbExclamation, mstrModule
        txtUserID.SetFocus
        'SendKeys "{Home}+{End}"
        Exit Sub
    End If
    '==============================================================================
        gintUserGroup = rst!UserGroup
        gstrUserID = rst!UserID
        gstrUserName = rst!UserName
        gstrUserPassword = rst!UserPassword
        gstrUserSalt = rst!Salt
        'gblnUserChangePassword = rst!ChangePassword
        If rst!Idle <> "" Then
            gintUserIdle = ConvInt(rst!Idle)
        End If
        gintUserIdle = 0 '10 ' Test
        mblnActive = rst!Active
        mintLoginAttempts = ConvInt(rst!LoginAttempts)
        strSalt = rst!Salt
        CloseRS rst
        CloseDB
        
        If gintUserIdle > 3600 Or gintUserIdle < 0 Then
            gintUserIdle = 0
        End If
        If mblnActive = False Then
            MsgBox "Your User ID has been frozen." & vbCrLf & _
                "Please contact System Administrator.", vbExclamation, mstrModule
            End
            Exit Sub
        End If
        '==============================================================================
        If mintLoginAttempts > 2 Then
            If mblnActive = True Then
                OpenDB
                UpdateField "UserData", "Active", "BOOLEAN", "TRUE", " UserID = '" & CheckInput(txtUserID.Text) & "'"
                CloseDB
            End If
            MsgBox "Your User ID has been frozen." & vbCrLf & _
                    "Please contact System Administrator.", vbExclamation, mstrModule
            End
            Exit Sub
        End If
        
        '==============================================================================
        'Encrypt password so we can check it against the encypted password in the database
        'Read in the salt
        strPassword = txtPassword.Text
        strPassword = strPassword & strSalt
    
        'Encrypt the entered password
        strPassword = GoldFishEncode(strPassword)
            
        If strPassword = gstrUserPassword Then
            'LoginSucceeded = True
            intAttempt = 0
            OpenDB
            UpdateField "UserData", "Loginattempts", "NUMBER", "0", " UserID = '" & CheckInput(txtUserID.Text) & "'"
            CloseDB
            'Me.Hide
            'Check if need to change password
            If NeedChangePassword(gstrUserID) = True Then
                Unload Me
                frmUserChangePassword.Show
                gblnUserChangePassword = False
                Exit Sub
            End If
            ' Check Module Access
            If UserAccessModule(MOD_DASHBOARD) Then
                Unload Me
                frmDashboard.Show
            Else
                MsgBox "Your access has been disabled!", vbExclamation, mstrModule
                Unload Me
            End If
            Exit Sub
        Else
            CloseRS rst
            'CloseDB
            If gintUserGroup > 1 Then
                intAttempt = intAttempt + 1
                OpenDB
                UpdateField "UserData", "Loginattempts", "NUMBER", "INCREMENT_ONE", "UserID = '" & CheckInput(txtUserID.Text) & "'"
                CloseDB
                If intAttempt > 2 Then
                    MsgBox "Too many attempts." & vbCrLf & _
                        "Application quit!", vbExclamation, mstrModule
                    End
                    'Exit Sub
                End If
            End If
        End If

    '==============================================================================
    MsgBox "Invalid Password, please try again!", vbExclamation, mstrModule
    'Show Form
    'Me.Show
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
CheckErr:
    Screen.MousePointer = vbDefault
    CloseRS rst
    CloseDB
    If Err.Number = 0 Then Exit Sub
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
    Const mstrMethod As String = "Form_Resize"
On Error GoTo CheckErr
    fraLogin.Left = (Me.Width - fraLogin.Width) \ 2
    fraLogin.Top = (Me.Height - fraLogin.Height) \ 2 '- fraLogin.Height
    'Line1.X1 = 150
    'Line1.X2 = Me.Width - 200
    'Line1.Y1 = fraLogin.Top + fraLogin.Height + 900 '- 5000
    'Line1.Y2 = Line1.Y1
    'lblCopyright.Left = Me.Width - 5000 'fraLogin.Left '300
    'lblCopyright.Width = Me.Width - 2000
    'lblCopyright.Top = Line1.Y1 + 100 'fraLogin.Top + fraLogin.Height + 900 'Me.Height - 4900
    Exit Sub
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogErrorText "Error", mstrMethod, Err.Description
    LogErrorDB "Sub", mstrModule, mstrMethod, Err.Number, Err.Description
End Sub

'Private Sub imgIcon_DblClick()
'    With frmUserChangePassword
'        Unload Me
'        .Show
'    End With
'End Sub

Private Sub lblCopyright_Click()
    txtUserID.Text = "ADMIN"
    txtPassword.Text = "admin"
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    KeyAscii = AscW(UCase$(ChrW$(KeyAscii)))
End Sub

Private Sub txtUserID_Validate(Cancel As Boolean)
    txtUserID.Text = UCase$(txtUserID.Text)
End Sub
