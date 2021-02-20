Attribute VB_Name = "modTextFile"
' Version : 2.1
'
' Modified On : 17/12/2014
' Descriptions : 1) Added sub Log2File
'
' Version : 2.0
' Modified On : 01/10/2014
' Descriptions : This Module is created by Aeric Poon
'                for handling Text file
Option Explicit

Public Sub Log2File(ByVal strPath As String, ByVal pstrText As String, _
                    Optional ByVal blnAddTimeStamp As Boolean, Optional ByVal blnAppend As Boolean = True)
    Dim FF As Integer
On Error GoTo WE
    FF = FreeFile
    If blnAppend Then
        Open strPath For Append As #FF
    Else
        Open strPath For Output As #FF
    End If
    If blnAddTimeStamp Then
        Print #FF, vbCrLf & FormatDateAndTime(Now) & vbCrLf & pstrText
    Else
        Print #FF, vbCrLf & pstrText
    End If
    Close #FF
    Exit Sub
WE:
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, App.Title
End Sub

Public Sub LogErrorText(FileName As String, pstrNote As String, Optional pstrError As String)
On Error GoTo SE
    Open App.Path & "\" & FileName & ".txt" For Append As #1
    If pstrError = "" Then
        Print #1, vbCrLf & FormatDateAndTime(Now) & vbCrLf & pstrNote
    Else
        Print #1, vbCrLf & FormatDateAndTime(Now) & vbCrLf & pstrNote & vbCrLf & pstrError
    End If
    Close #1
    Exit Sub
SE:
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbExclamation, App.Title
End Sub

Public Function WriteTextFile(pstrFileFullPath As String, pstrText As String)
    Dim FF As Integer
On Error GoTo WE
    FF = FreeFile
    Open pstrFileFullPath For Output As #FF
        Print #FF, CStr(pstrText)
    Close #FF
    Exit Function
WE:
    LogErrorText "Error", "WriteTextFile(" & pstrFileFullPath & ")", Err.Description
End Function

Public Sub ReadTextFile(ByVal FileName As String, ByVal LineNo As Integer, ByRef sOutput As String)
    Dim FF As Integer
    Dim i As Integer
On Error GoTo RE
    FF = FreeFile
    Open App.Path & "\" & FileName & ".txt" For Input As #FF
        If LineNo < 0 Then
            Do Until EOF(FF) = True
                Input #FF, sOutput
            Loop
        ElseIf LineNo > 0 Then
            For i = 0 To LineNo
                If Not EOF(FF) Then
                    Input #FF, sOutput
                Else
                    Exit For
                End If
            Next
        Else
            Input #FF, sOutput
        End If
    Close
    Exit Sub
RE:
    LogErrorText "Error", "ReadTextFile(" & FileName & ", " & LineNo & ")", Err.Description
End Sub

Public Sub ReadFile(ByVal strPath As String, ByVal LineNo As Integer, ByRef strOutput As String)
    Dim FF As Integer
    Dim i As Integer
On Error GoTo RE
    FF = FreeFile
    Open strPath For Input As #FF
        If LineNo < 0 Then
            Do Until EOF(FF) = True
                Input #FF, strOutput
            Loop
        ElseIf LineNo > 0 Then
            For i = 0 To LineNo
                If Not EOF(FF) Then
                    Input #FF, strOutput
                Else
                    Exit For
                End If
            Next
        Else
            Input #FF, strOutput
        End If
    Close
    Exit Sub
RE:
    LogErrorText "Error", "ReadFile(" & strPath & ", " & LineNo & ")", Err.Description
End Sub

Public Function WriteFile(ByVal strPath As String, ByVal strText As String)
    Dim FF As Integer
On Error GoTo WE
    FF = FreeFile
    Open strPath For Output As #FF
        Print #FF, CStr(strText)
    Close #FF
    Exit Function
WE:
    LogErrorText "Error", "WriteFile(" & strPath & ")", Err.Description
End Function
