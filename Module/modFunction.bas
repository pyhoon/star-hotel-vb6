Attribute VB_Name = "modFunction"
' Version : 2.1
' Modified On : 04/01/2015
' Descriptions : 1) Modify function UserAccessModule to allow check for other UserID

' Modified On : 01/10/2014
' Descriptions : 1) Add function LogErrorDB

Option Explicit
Private Const mstrModule As String = "modFunction"

' Get Database Version
Public Function DB_Version() As Double
    Const mstrMethod As String = "DB_Version"
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    Dim dblVersion As String
On Error GoTo CheckErr
    strSQL = "SELECT DatabaseVersion"
    strSQL = strSQL & " FROM Company"
    OpenDB
    Set rst = OpenRS(strSQL)
    If Not (rst.EOF Or rst.BOF) Then
        dblVersion = Val(rst!DatabaseVersion)
    End If
    DB_Version = dblVersion
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

'Public Sub DB_Version_Update(strVersion As String)
'    Const mstrMethod As String = "DB_Version_Update"
'    Dim strSQL As String
'On Error GoTo CheckErr
'    If strVersion = "" Then Exit Sub
'    strSQL = "UPDATE Company"
'    strSQL = strSQL & " SET DatabaseVersion = " & strVersion
'    OpenDB
'    QuerySQL strSQL
'    CloseDB
'    Exit Sub
'CheckErr:
'    CloseDB
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
'    'LogError "Error", mstrMethod, Err.Description
'    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
'End Sub

Public Function UserGroup(pstrUserID As String) As Long
    Const mstrMethod As String = "UserGroup"
    Dim rst As ADODB.Recordset
    Dim strSQL As String
On Error GoTo CheckErr
    strSQL = "SELECT UserGroup"
    strSQL = strSQL & " FROM UserData"
    strSQL = strSQL & " WHERE UserID = '" & pstrUserID & "'"
    OpenDB
    Set rst = OpenRS(strSQL)
    If Not rst.EOF Then
        UserGroup = CLng(rst(0).Value)
    Else
        UserGroup = 0
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    UserGroup = 0
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function UserAccessModule(intModuleID As Integer, Optional strUserID As String = "") As Boolean
    Const mstrMethod As String = "UserAccessModule"
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    Dim blnGroup(1 To 4) As Boolean
On Error GoTo CheckErr
    If strUserID = "" Then strUserID = gstrUserID
    strSQL = "SELECT * FROM ModuleAccess"
    strSQL = strSQL & " WHERE ModuleID = " & intModuleID
    OpenDB
    Set rst = OpenRS(strSQL)
    If Not rst.EOF Then
        blnGroup(1) = rst!Group1
        blnGroup(2) = rst!Group2
        blnGroup(3) = rst!Group3
        blnGroup(4) = rst!Group4
    Else
        blnGroup(1) = False
        blnGroup(2) = False
        blnGroup(3) = False
        blnGroup(4) = False
    End If
    CloseRS rst
    strSQL = "SELECT UserGroup"
    strSQL = strSQL & " FROM UserData"
    strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
    Set rst = OpenRS(strSQL)
    If Not rst.EOF Then
        If rst!UserGroup = 1 Then
            UserAccessModule = blnGroup(1)
        ElseIf rst!UserGroup = 2 Then
            UserAccessModule = blnGroup(2)
        ElseIf rst!UserGroup = 3 Then
            UserAccessModule = blnGroup(3)
        Else ' rst!UserGroup = 4
            UserAccessModule = blnGroup(4)
        End If
    Else
        UserAccessModule = False
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    UserAccessModule = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function NeedChangePassword(strUserID As String) As Boolean
    Const mstrMethod As String = "NeedChangePassword"
    Dim rst As ADODB.Recordset
    Dim strSQL As String
On Error GoTo CheckErr
    strSQL = "SELECT ChangePassword"
    strSQL = strSQL & " FROM UserData"
    strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
    OpenDB
    Set rst = OpenRS(strSQL)
    If Not rst.EOF Then
        If rst!ChangePassword = True Then
            NeedChangePassword = True
        Else
            NeedChangePassword = False
        End If
    Else
        NeedChangePassword = False
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    NeedChangePassword = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function AdminUser(strUserID As String) As Boolean
    Const mstrMethod As String = "AdminUser"
    Dim rst As ADODB.Recordset
    Dim strSQL As String
On Error GoTo CheckErr
    strSQL = "SELECT UserGroup"
    strSQL = strSQL & " FROM UserData"
    strSQL = strSQL & " WHERE UserID = '" & strUserID & "'"
    OpenDB
    Set rst = OpenRS(strSQL)
    If Not rst.EOF Then
        If rst!UserGroup = 1 Then
            AdminUser = True
        Else
            AdminUser = False
        End If
    Else
        AdminUser = False
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    AdminUser = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function QueryHasData(pstrQuery As String) As Boolean
    Const mstrMethod As String = "QueryHasData"
    Dim rst As ADODB.Recordset
On Error GoTo CheckErr
    OpenDB
    Set rst = OpenRS(pstrQuery)
    If Not rst.RecordCount = 0 Then
        QueryHasData = True
    Else
        QueryHasData = False
    End If
    CloseRS rst
    CloseDB
    Exit Function
CheckErr:
    CloseRS rst
    CloseDB
    QueryHasData = False
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

'Quick way to List records from Database table
Public Function GetList(pstrTable As String, _
                        pstrColumnID As String, _
                        pstrColumnName As String, _
                        Optional pstrCondition As String, _
                        Optional pstrColumnOrder As String, _
                        Optional pblnSortAscending As Boolean = True) As ADODB.Recordset
    Const mstrMethod As String = "GetList"
    Dim rst As ADODB.Recordset
    Dim strSQL As String
On Error GoTo CheckErr
    strSQL = "SELECT " & pstrColumnID & ","
    strSQL = strSQL & pstrColumnName
    If pstrColumnOrder <> "" Then
        strSQL = strSQL & ", " & pstrColumnOrder
    End If
    strSQL = strSQL & " FROM " & pstrTable
    If pstrCondition <> "" Then
        strSQL = strSQL & " WHERE " & pstrCondition
    End If
    If pstrColumnOrder <> "" Then
        strSQL = strSQL & " ORDER BY " & pstrColumnOrder
        If Not pblnSortAscending Then
            strSQL = strSQL & " DESC"
        End If
    End If
    Set rst = OpenRS(strSQL)
    'Debug.Print strSQL
    Set GetList = rst
    Exit Function
CheckErr:
    CloseRS rst
    Set GetList = Nothing
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

'Quick way to Select One field in Database table
Public Function SelectField(pstrTable As String, pstrColumn As String, Optional pstrCondition As String) As String
    Const mstrMethod As String = "SelectField"
    Dim rst As ADODB.Recordset
    Dim strSQL As String
On Error GoTo CheckErr
    strSQL = "SELECT " & pstrColumn & " FROM " & pstrTable
    If pstrCondition <> "" Then strSQL = strSQL & " WHERE " & pstrCondition
    Set rst = OpenRS(strSQL)
    SelectField = rst(0).Value
    CloseRS rst
    Exit Function
CheckErr:
    CloseRS rst
    SelectField = ""
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

'Quick way to Update One field in Database table
Public Function UpdateField(pstrTable As String, pstrColumn As String, pstrType As String, pstrValue As String, Optional pstrCondition As String) As ADODB.Recordset
    Const mstrMethod As String = "UpdateField"
    Dim strSQL As String
On Error GoTo CheckErr
    strSQL = "UPDATE " & pstrTable & " SET " & pstrColumn & " = "
    If UCase(pstrType) = "STRING" Then
        strSQL = strSQL & "'" & pstrValue & "'"
    Else
        If pstrValue = "INCREMENT_ONE" Then
            strSQL = strSQL & pstrColumn & " + 1"
        Else
            strSQL = strSQL & pstrValue
        End If
    End If
    If pstrCondition <> "" Then strSQL = strSQL & " WHERE " & pstrCondition
    With ACN
        .BeginTrans
        .Execute strSQL
        .CommitTrans
    End With
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    'LogError "Error", mstrMethod, Err.Description
    LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Sub LogErrorDB(LogType As String, LogModule As String, LogMethod As String, ErrorNumber As String, Optional ErrorDescription As String)
    Const mstrMethod As String = "LogErrorDB"
    Dim strSQL As String
On Error GoTo CheckErr
    If LogType = "" Then LogType = "Unknown"
    strSQL = "INSERT INTO LogError ("
    strSQL = strSQL & " LogDateTime,"
    strSQL = strSQL & " LogErrorNum,"
    strSQL = strSQL & " LogErrorDescription,"
    strSQL = strSQL & " LogUserName,"
    strSQL = strSQL & " LogModule,"
    strSQL = strSQL & " LogMethod,"
    strSQL = strSQL & " LogType)"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & "#" & FormatDateAndTime(Now) & "#,"
    strSQL = strSQL & " '" & CheckInput(ErrorNumber) & "',"
    strSQL = strSQL & " '" & CheckInput(ErrorDescription) & "',"
    strSQL = strSQL & " '" & gstrUserID & "',"
    strSQL = strSQL & " '" & CheckInput(LogModule) & "',"
    strSQL = strSQL & " '" & CheckInput(LogMethod) & "',"
    strSQL = strSQL & " '" & CheckInput(LogType) & "')"
    OpenDB
    With ACN
        .BeginTrans
        .Execute strSQL
        .CommitTrans
    End With
    CloseDB
    Exit Sub
CheckErr:
    'MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
End Sub
