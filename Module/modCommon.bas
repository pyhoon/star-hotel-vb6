Attribute VB_Name = "modCommon"
' Version : 2.2.3
'
' Modified On : 19/12/2014
' Descriptions : 1) Added function ConvText
'
' Modified On : 02/11/2014
' Descriptions : 1) Added function ConvInt
'
' Modified On : 28/10/2014
' Descriptions : 1) Modify SQL_INNER_JOIN by removing parameters
'                2) Modify SQL_LEFT_JOIN by removing parameters
'                3) Add new sub SQL_ON
'
' Modified On : 03/10/2014
' Descriptions : 1) Remove SQL comma at in front
'                2) Rename SQL Functions Name
'                3) Fixed some Date Functions, Add MonthDay30
'                4) Add Optional blnBeginSpace As Boolean = True
'                5) Add function SQL_SET_Integer
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
'Private Const CP_UTF8   As Long = 65001

Public Function CheckString(strInput As String) As String
    'strInput = Replace(strInput, "'", "")
    strInput = Replace(strInput, "'", "''")
    strInput = Replace(strInput, """", "")
    ' Error in SQL if use following in LIKE search
    strInput = Replace(strInput, "[", "")
    strInput = Replace(strInput, "]", "")
    strInput = Replace(strInput, "*", "")
    strInput = Replace(strInput, "&", "")
    CheckString = strInput
End Function

' Prevent Hacker's SQL Injection
' Try to key in 1=1'or'1=1 as password to hack
Public Function CheckInput(strInput As String) As String
    'Take out parts of the username that are not permitted
    strInput = Replace(strInput, "password", "", 1, -1, 1)
    strInput = Replace(strInput, "salt", "", 1, -1, 1)
    strInput = Replace(strInput, "author", "", 1, -1, 1)
    strInput = Replace(strInput, "code", "", 1, -1, 1)
    strInput = Replace(strInput, "username", "", 1, -1, 1)
    strInput = Replace(strInput, "select", "", 1, -1, 1)
    strInput = Replace(strInput, "from", "", 1, -1, 1)
    'Replace harmful SQL quotation marks with doubles
    'CheckInput = strInput 'Test
    strInput = Replace(strInput, """", "", 1, -1, 1) 'Use this
    strInput = Replace(strInput, "'", "''", 1, -1, 1)   'Use this
    'strInput = Replace(strInput, "''", " ", 1, -1, 1) 'Do not use this
    CheckInput = strInput
End Function

Public Function ConvText(pvarInput As Variant) As String
    If Trim(pvarInput) <> "" Then
        ConvText = Trim(pvarInput)
    Else
        ConvText = ""
    End If
End Function

Public Function ConvInt(pstrInput As String) As Integer
    If Trim(pstrInput) = "" Then
        ConvInt = 0
    Else
        If IsNumeric(pstrInput) = True Then
            ConvInt = CInt(pstrInput)
        Else
            ConvInt = 0
        End If
    End If
End Function

Public Function ConvLng(pstrInput As String) As Long
    If Trim(pstrInput) = "" Then
        ConvLng = 0
    Else
        If IsNumeric(pstrInput) = True Then
            ConvLng = CLng(pstrInput)
        Else
            ConvLng = 0
        End If
    End If
End Function

Public Function ConvDbl(pstrInput As String) As Double
    If Trim(pstrInput) = "" Then
        ConvDbl = 0
    ElseIf IsNull(pstrInput) Then
        ConvDbl = 0
    Else
        If IsNumeric(pstrInput) = True Then
            ConvDbl = CDbl(pstrInput)
        Else
            ConvDbl = 0
        End If
    End If
End Function

Public Function ConvCur(pstrInput As String) As Currency
    If Trim(pstrInput) = "" Then
        ConvCur = 0
    ElseIf IsNull(pstrInput) Then
        ConvCur = 0
    Else
        If IsNumeric(pstrInput) = True Then
            ConvCur = CCur(pstrInput)
        Else
            ConvCur = 0
        End If
    End If
End Function

Public Function FormatCurrency(pstrInput As Variant, Optional intDecimal As Integer = 2) As String
    FormatCurrency = FormatNumber(pstrInput, intDecimal, vbTrue, vbFalse, vbTrue)
End Function

Public Function FormatDate(dtDate As Date) As String
    FormatDate = Format(dtDate, "dd MMM yyyy")
End Function

Public Function FormatTime(dtDate As Date) As String
    FormatTime = Format(dtDate, "hh:mm AMPM")
End Function

Public Function FormatTimeSeconds(dtDate As Date, Optional strAMPM As String = " AMPM") As String
    FormatTimeSeconds = Format(dtDate, "hh:mm:ss" & strAMPM)
End Function

Public Function FormatDateAndTime(dtDate As Date) As String
    FormatDateAndTime = Format(dtDate, "dd MMM yyyy hh:mm:ss AMPM")
End Function

Public Function FormatDateDayAndTime(dtDate As Date) As String
    FormatDateDayAndTime = Format(dtDate, "DDDD, dd MMMM yyyy     hh:mm AMPM")
End Function

Public Function FormatMonthYear(dtDate As Date) As String
    FormatMonthYear = Format(dtDate, "MMM yyyy")
End Function

Public Function FormatYear(dtDate As Date) As String
    FormatYear = Format(dtDate, "yyyy")
End Function

Public Function WeekDay1(dtDate As Date) As String
    Dim intDay As Integer
    Dim datDate As Date
    intDay = Weekday(dtDate, vbSunday)
    datDate = DateAdd("D", -intDay + 1, dtDate)
    WeekDay1 = FormatDate(datDate)
End Function

Public Function WeekDayN(dtDate As Date, intNthDay As Integer) As String
    Dim intDay As Integer
    Dim datDate As Date
    intDay = Weekday(dtDate, vbSunday)
    datDate = DateAdd("D", intNthDay - intDay, dtDate)
    WeekDayN = FormatDate(datDate)
End Function

Public Function WeekDay7(dtDate As Date) As String
    Dim intDay As Integer
    Dim datDate As Date
    intDay = Weekday(dtDate, vbSunday)
    datDate = DateAdd("D", 7 - intDay, dtDate)
    WeekDay7 = FormatDate(datDate)
End Function

Public Function MonthDay1(dtDate As Date) As String
    Dim strTemp As String
    Dim datDate As Date
    strTemp = "1 "
    strTemp = strTemp & MonthName(Month(dtDate), True) & " "
    strTemp = strTemp & Year(dtDate)
    datDate = DateAdd("D", 0, strTemp)
    MonthDay1 = FormatDate(datDate)
End Function

Public Function MonthDay30(dtDate As Date) As String
    Dim strTemp As String
    Dim datDate As Date
    Dim intNextMonth As Integer
    Dim intNextMonthYear As Integer
    If Month(dtDate) = 12 Then
        intNextMonth = 1
        intNextMonthYear = Year(dtDate) + 1
    Else
        intNextMonth = Month(dtDate) + 1
        intNextMonthYear = Year(dtDate)
    End If
    strTemp = "1 " & MonthName(intNextMonth, True) & " " & intNextMonthYear
    datDate = DateAdd("D", -1, strTemp)
    MonthDay30 = FormatDate(datDate)
End Function

Public Function YearDay1(dtDate As Date) As String
    Dim strTemp As String
    Dim datDate As Date
    strTemp = "1 "
    strTemp = strTemp & MonthName(1, True) & " "
    strTemp = strTemp & Year(dtDate)
    datDate = DateAdd("D", 0, strTemp)
    YearDay1 = FormatDate(datDate)
End Function

Public Function YearDay365(dtDate As Date) As String
    Dim strTemp As String
    Dim datDate As Date
    strTemp = "31 "
    strTemp = strTemp & MonthName(12, True) & " "
    strTemp = strTemp & Year(dtDate)
    datDate = DateAdd("D", 0, strTemp)
    YearDay365 = FormatDate(datDate)
End Function

Public Function FormatDigit(lngID As Long, intDigit As Integer) As String
    Dim i As Integer
    Dim strTemp As String
    Dim intLen As Integer
    intLen = Len(CStr(lngID))
    For i = 1 To (intDigit - intLen)
        strTemp = strTemp & "0"
    Next
    FormatDigit = strTemp & lngID 'Format$(lngTransID, "0000000")
End Function

Public Function FormatFixedID(strID As String, intLength As Integer) As String
    Dim i As Integer
    FormatFixedID = strID
    For i = 1 To (intLength - Len(strID))
        FormatFixedID = FormatFixedID & " "
    Next
End Function

Public Function FormatItem(strItemID As String, strItemName As String, Optional blnActive As Boolean = True) As String
    Dim str As String
    str = "(X)" ' "Ø"
    If blnActive = True Then
        FormatItem = strItemID & vbTab & strItemName
    Else
        FormatItem = strItemID & str & vbTab & strItemName
    End If
End Function

Public Function GetItemID(strItem As String) As String
    Dim strPart() As String
    strPart = Split(strItem, vbTab, 2)
    strItem = Replace(strPart(0), "*", "")
    GetItemID = strItem
End Function

'Public Sub SetString(txtOutput As TextBox, adoField As ADODB.Field, Optional blnTrim As Boolean = False, Optional strDefault As String)
'    If adoField <> strDefault Then
'        If blnTrim Then
'            txtOutput.Text = Trim(adoField)
'        Else
'            txtOutput.Text = adoField
'        End If
'    Else
'        txtOutput.Text = strDefault
'    End If
'End Sub

'Public Sub SetNumeric(txtOutput As TextBox, adoField As ADODB.Field, Optional strType As String = "", Optional strDefault As String = "0")
'    If adoField <> strDefault Then
'        If blnTrim Then
'            txtOutput.Text = Trim(adoField)
'        Else
'            txtOutput.Text = adoField
'        End If
'    Else
'        txtOutput.Text = strDefault
'    End If
'End Sub

Public Sub SetCombo(cboOutput As ComboBox, strText As String, Optional intItemData As Integer = 0)
    Dim i As Integer
    If intItemData <> 0 Then
        For i = 0 To cboOutput.ListCount - 1
            If cboOutput.ItemData(i) = intItemData Then
                cboOutput.ListIndex = i
                Exit For
            End If
        Next
    Else
        If strText <> "" Then
            For i = 0 To cboOutput.ListCount - 1
                If cboOutput.List(i) = strText Then
                    cboOutput.ListIndex = i
                    Exit For
                End If
            Next
        Else
            cboOutput.ListIndex = 0
        End If
    End If
End Sub

Public Sub SetRecord(txtOutput As TextBox, adoField As ADODB.Field, Optional blnTrim As Boolean = True)
    If adoField.Value <> "" Then
        If blnTrim = True Then
            txtOutput.Text = Trim(adoField)
        Else
            txtOutput.Text = adoField
        End If
    Else
        txtOutput.Text = ""
    End If
End Sub

Public Sub SetCheck(chkOutput As CheckBox, adoField As ADODB.Field)
    If adoField = True Then
        chkOutput.Value = vbChecked
    Else
        chkOutput.Value = vbUnchecked
    End If
End Sub

Public Sub SQL_SET_Text(strField As String, strText As String, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = '" & strText & "'"
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_Double(strField As String, dblNumber As Double, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = " & dblNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_Integer(strField As String, intNumber As Integer, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = " & intNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_Long(strField As String, lngNumber As Long, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = " & lngNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_Boolean(strField As String, blnValue As Boolean, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField
    If blnValue Then
        gstrSQL = gstrSQL & " = TRUE"
    Else
        gstrSQL = gstrSQL & " = FALSE"
    End If
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_DateTime(strField As String, strDateTime As String, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = #" & strDateTime & "#"
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLText(strText As String, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        gstrSQL = gstrSQL & " "
    End If
    gstrSQL = gstrSQL & strText
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_WHERE_Text(strField As String, strText As String)
    gstrSQL = gstrSQL & " WHERE " & strField & " = '" & strText & "'"
End Sub

Public Sub SQL_WHERE_Long(strField As String, lngNumber As Long)
    gstrSQL = gstrSQL & " WHERE " & strField & " = " & lngNumber
End Sub

Public Sub SQL_WHERE_Integer(strField As String, intNumber As Integer)
    gstrSQL = gstrSQL & " WHERE " & strField & " = " & intNumber
End Sub

Public Sub SQL_WHERE_Boolean(strField As String, blnBoolean As Boolean)
    gstrSQL = gstrSQL & " WHERE " & strField & " = " & blnBoolean
End Sub

Public Sub SQL_WHERE_BETWEEN(strField As String, strLeftValue As String, strRightValue As String)
    gstrSQL = gstrSQL & " WHERE " & strField & " BETWEEN " & strLeftValue & " AND " & strRightValue
End Sub

Public Sub SQL_WHERE_LIKE_Text(strField As String, strText As String)
    gstrSQL = gstrSQL & " WHERE " & strField & " LIKE '%" & strText & "%'"
End Sub

Public Sub SQL_AND_LIKE_Text(strField As String, strText As String)
    gstrSQL = gstrSQL & " AND " & strField & " LIKE '%" & strText & "%'"
End Sub

Public Sub SQL_OR_LIKE_Text(strField As String, strText As String)
    gstrSQL = gstrSQL & " OR " & strField & " LIKE '%" & strText & "%'"
End Sub

Public Sub SQL_ORDER_BY(strField As String, Optional blnAscending As Boolean = True)
    gstrSQL = gstrSQL & " ORDER BY " & strField
    If blnAscending = False Then
        gstrSQL = gstrSQL & " DESC"
    End If
End Sub

'Public Sub SQL_INNER_JOIN(strTable1 As String, strTable2 As String, strCommonField1 As String, strCommonField2 As String)
'    'SQLText "FROM " & strTable1, False
'    SQLText "INNER JOIN " & strTable2, False
'    SQLText "ON " & strTable1 & "." & strCommonField1 & " = " & strTable2 & "." & strCommonField2, False
'End Sub

'Public Sub SQL_INNER_JOIN(strTable1 As String, strTable2 As String, strCommonField1 As String, strCommonField2 As String, Optional strAlias1 As String = "", Optional strAlias2 As String = "")
'    'SQLText "FROM " & strTable1, False
'    If strAlias1 = "" Then
'        strAlias1 = strTable1
'    End If
'    If strAlias2 = "" Then
'        strAlias2 = strTable2
'    End If
'    SQLText "INNER JOIN " & strTable2, False
'    If strAlias2 <> "" Then
'        SQLText strAlias2, False
'    End If
'    SQLText "ON " & strAlias1 & "." & strCommonField1 & " = " & strAlias2 & "." & strCommonField2, False
'End Sub

Public Sub SQL_INNER_JOIN(strTable2 As String, Optional strAlias2 As String = "")
    SQLText "INNER JOIN " & strTable2, False
    If strAlias2 <> "" Then
        SQLText strAlias2, False
    End If
End Sub

Public Sub SQL_LEFT_JOIN(strTable2 As String, Optional strAlias2 As String = "")
    SQLText "LEFT JOIN " & strTable2, False
    If strAlias2 <> "" Then
        SQLText strAlias2, False
    End If
End Sub

Public Sub SQL_ON(strAlias1 As String, strCommonField1 As String, strAlias2 As String, strCommonField2 As String)
    SQLText "ON " & strAlias1 & "." & strCommonField1 & " = " & strAlias2 & "." & strCommonField2, False
End Sub

Public Sub SQL_SELECT()
    gstrSQL = "SELECT"
End Sub

Public Sub SQL_SELECT_ALL(strTable As String)
    gstrSQL = "SELECT * FROM " & strTable
End Sub

Public Sub SQL_SELECT_TOP(strField As String, strTable As String, Optional intTop As Integer = 1)
    gstrSQL = "SELECT TOP " & intTop & " " & strField & " FROM " & strTable
End Sub

Public Sub SQL_FROM(strTable As String, Optional strAlias As String = "")
    If strAlias <> "" Then
        gstrSQL = gstrSQL & " FROM " & strTable & " " & strAlias
    Else
        gstrSQL = gstrSQL & " FROM " & strTable
    End If
End Sub

Public Sub SQL_INSERT(strTable As String)
    gstrSQL = "INSERT INTO " & strTable & " ("
End Sub

Public Sub SQL_VALUES()
    gstrSQL = gstrSQL & ") VALUES ("
End Sub

Public Sub SQL_UPDATE(strTable As String)
    gstrSQL = "UPDATE " & strTable & " SET"
End Sub

Public Sub SQL_DELETE(strTable As String)
    gstrSQL = "DELETE FROM " & strTable
End Sub

Public Sub SQL_DROP(strTable As String)
    gstrSQL = "DROP TABLE [" & strTable & "]"
End Sub

Public Sub SQL_ALTER_TABLE(strTable As String)
    gstrSQL = "ALTER TABLE " & strTable
End Sub

Public Sub SQL_CREATE(strTable As String, Optional strPrefix As String = "")
    gstrSQL = "CREATE TABLE " & strPrefix & strTable
    gstrSQL = gstrSQL & " ("
End Sub

Public Sub SQL_COLUMN_ID(Optional strColumnName As String = "ID", Optional blnPrimaryKey As Boolean = True, Optional blnAutoIncrement As Boolean = True, Optional blnEndComma As Boolean = True)
    'gstrSQL = gstrSQL & "[" & strColumnName & "]"
    gstrSQL = gstrSQL & strColumnName
    If blnAutoIncrement Then
        gstrSQL = gstrSQL & " AUTOINCREMENT"
    Else
        gstrSQL = gstrSQL & " LONG"
    End If
    If blnPrimaryKey Then gstrSQL = gstrSQL & " PRIMARY KEY"
    If blnEndComma = True Then gstrSQL = gstrSQL & "," 'gstrSQL = gstrSQL & ","
End Sub

' Short Text
Public Sub SQL_COLUMN_TEXT(strColumnName As String, Optional intLength As Integer = 255, Optional strDefault As String = "", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then gstrSQL = gstrSQL & " "
    gstrSQL = gstrSQL & strColumnName & " TEXT(" & intLength & ")"
    If strDefault <> "" Then gstrSQL = gstrSQL & " DEFAULT """ & strDefault & """"
    If Not blnNullable Then gstrSQL = gstrSQL & " NOT NULL"
    If blnEndComma = True Then gstrSQL = gstrSQL & ","
End Sub

' Long Text
Public Sub SQL_COLUMN_MEMO(strColumnName As String, Optional strDefault As String = "", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then gstrSQL = gstrSQL & " "
    gstrSQL = gstrSQL & strColumnName & " MEMO"
    If strDefault <> "" Then gstrSQL = gstrSQL & " DEFAULT " & strDefault
    If Not blnNullable Then gstrSQL = gstrSQL & " NOT NULL"
    If blnEndComma = True Then gstrSQL = gstrSQL & ","
End Sub

' NOTE: Not yet used or tested
Public Sub SQL_COLUMN_NUMBER(strColumnName As String, Optional strFieldSize As String = "LONG", Optional strDefault As String = "", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then gstrSQL = gstrSQL & " "
    Select Case strFieldSize
        Case "BYTE"
            gstrSQL = gstrSQL & strColumnName & " BYTE"
        Case "SHORT"
            gstrSQL = gstrSQL & strColumnName & " SHORT"
        Case "INTEGER" ' Same as SHORT ?
            gstrSQL = gstrSQL & strColumnName & " INTEGER"
        Case "LONG" ' Default
            gstrSQL = gstrSQL & strColumnName & " LONG"
        Case "SINGLE"
            gstrSQL = gstrSQL & strColumnName & " SINGLE"
        Case "DOUBLE"
            gstrSQL = gstrSQL & strColumnName & " DOUBLE"
        Case "CURRENCY"
            gstrSQL = gstrSQL & strColumnName & " CURRENCY"
        Case "REPLICA", "GUID"
            gstrSQL = gstrSQL & strColumnName & " GUID"
        Case "DECIMAL"
            gstrSQL = gstrSQL & strColumnName & " DECIMAL (18, 0)" ' (precision, scale) 9, 4
        Case Else ' LONG
            gstrSQL = gstrSQL & strColumnName & " LONG"
    End Select
    If strDefault <> "" Then gstrSQL = gstrSQL & " DEFAULT " & strDefault
    If Not blnNullable Then gstrSQL = gstrSQL & " NOT NULL"
    If blnEndComma = True Then gstrSQL = gstrSQL & ","
End Sub

Public Sub SQL_COLUMN_BIT(strColumnName As String, Optional strDefault As String = "-1", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then gstrSQL = gstrSQL & " "
    gstrSQL = gstrSQL & strColumnName & " BIT"
    If strDefault <> "" Then gstrSQL = gstrSQL & " DEFAULT " & strDefault
    If Not blnNullable Then gstrSQL = gstrSQL & " NOT NULL"
    If blnEndComma = True Then gstrSQL = gstrSQL & ","
End Sub

' Same as SQL_COLUMN_BIT
Public Sub SQL_COLUMN_YESNO(strColumnName As String, Optional strDefault As String = "Yes", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then gstrSQL = gstrSQL & " "
    gstrSQL = gstrSQL & strColumnName & " YESNO"
    If strDefault <> "" Then gstrSQL = gstrSQL & " DEFAULT " & strDefault
    If Not blnNullable Then gstrSQL = gstrSQL & " NOT NULL"
    If blnEndComma = True Then gstrSQL = gstrSQL & ","
End Sub

Public Sub SQL_COLUMN_DATETIME(strColumnName As String, Optional strDefault As String = "", Optional blnNullable As Boolean = True, Optional blnBeginSpace As Boolean = True, Optional blnEndComma As Boolean = True)
    If blnBeginSpace Then gstrSQL = gstrSQL & " "
    gstrSQL = gstrSQL & strColumnName & " DATETIME"
    If strDefault <> "" Then gstrSQL = gstrSQL & " DEFAULT " & strDefault
    If Not blnNullable Then gstrSQL = gstrSQL & " NOT NULL"
    If blnEndComma = True Then gstrSQL = gstrSQL & ","
End Sub

Public Sub SQL_Comma()
    gstrSQL = gstrSQL & ","
End Sub

Public Sub SQL_Close_Bracket()
    gstrSQL = gstrSQL & ")"
End Sub

Public Sub SQLData_Text(strText As String, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        gstrSQL = gstrSQL & " "
    End If
    gstrSQL = gstrSQL & "'" & CheckString(strText) & "'"
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_Double(dblNumber As Double, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    gstrSQL = gstrSQL & " " & dblNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_Long(lngNumber As Long, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    gstrSQL = gstrSQL & " " & lngNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_Integer(intNumber As Integer, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    gstrSQL = gstrSQL & " " & intNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_Boolean(blnValue As Boolean, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnValue Then
        gstrSQL = gstrSQL & " TRUE"
    Else
        gstrSQL = gstrSQL & " FALSE"
    End If
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_DateTime(strDateTime As String, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    gstrSQL = gstrSQL & " #" & strDateTime & "#"
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Function GenWord()
    Dim strPassword As String
    Dim intArray(10) As Integer
    Dim l As Integer

    intArray(0) = 67
    intArray(1) = 111
    intArray(2) = 109
    intArray(3) = 112
    intArray(4) = 117
    intArray(5) = 116
    intArray(6) = 101
    intArray(7) = 114
    intArray(8) = 105
    intArray(9) = 115
    intArray(10) = 101

    strPassword = ""
    For l = 0 To UBound(intArray)
        strPassword = strPassword & Chr(intArray(l))
    Next
    GenWord = strPassword
End Function

Public Sub WriteText(FileName As String, Optional sNote As String)
    Open App.Path & "\" & FileName & ".txt" For Append As #1
    If sNote = "" Then
        Write #1, FormatDateAndTime(Now), Error
    Else
        Write #1, FormatDateAndTime(Now), Error & " @ " & sNote
    End If
    Close #1
End Sub

Public Sub ReadText(ByVal FileName As String, ByVal LineNo As Integer, ByRef sOutput As String)
    On Error GoTo NewFile
    Dim i As Integer
    Open App.Path & "\" & FileName & ".txt" For Input As #2
        If LineNo < 0 Then
            Do Until EOF(2) = True
                Input #2, sOutput
            Loop
        ElseIf LineNo > 0 Then
            For i = 0 To LineNo
                If Not EOF(2) Then
                    Input #2, sOutput
                Else
                    Exit For
                End If
            Next
        Else
            Input #2, sOutput
        End If
    Close
    Exit Sub
NewFile:
    WriteText "Error", "ReadText(" & FileName & ".txt)"
End Sub

Public Function FileExists(strPath As String) As Boolean
    Dim lngRetVal As Long
    On Error Resume Next
    lngRetVal = Len(Dir$(strPath))
    If Err Or lngRetVal = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Public Sub ResetControls(frm As Form)
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then ctl.Text = vbNullString
        If TypeOf ctl Is ComboBox Then ctl.ListIndex = -1
    Next
End Sub
