Attribute VB_Name = "modEncryption"
' Version : 1.0
' Modified On : 01/10/2014
' Descriptions : This Module is created by Aeric Poon for simple encryption
Option Explicit
Private Const mstrModule As String = "modEncryption"

Public Function Encrypt(strPlaintext As String, strSalt As String) As String
Const mstrMethod As String = "Encrypt"
On Error GoTo CheckErr
    'GoldFishEncode with a random Salt
    Encrypt = GoldFishEncode(strPlaintext & strSalt)
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

Public Function GenSalt(StringLen As Integer) As String
Const mstrMethod As String = "GenSalt"
On Error GoTo CheckErr
    'Temporary make as simple as possible
    Dim i As Integer
    'Max length = 6
    If StringLen > 6 Then StringLen = 6
    StringLen = StringLen \ 2
    For i = 1 To StringLen
        Randomize
        GenSalt = GenSalt & Hex((Rnd() * 64) Mod 100)
    Next
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

'GoldFish Encoding
Public Function GoldFishEncode(pw As String)
Const mstrMethod As String = "GoldFishEncode"
Dim M As Integer
Dim i As Integer
Dim B As Integer
Dim ibin As String
Dim xbin As String
Dim ybin As String
Dim d As Integer
Dim n As Integer
On Error GoTo CheckErr
    ibin = Dec2Bin(Len(pw))
    For M = 1 To Len(pw)
        xbin = Dec2Bin(Asc(Mid(pw, M, 1)))
        ybin = ""
        For i = 1 To 8
            ybin = ybin & (CInt(Mid(xbin, i, 1)) + CInt(Mid(ibin, i, 1))) Mod 2
        Next
        d = 8
        n = 0
        For i = 1 To 8
            B = Mid(ybin, i, 1) * 2 ^ (d - 1)
            n = n + B
            d = d - 1
        Next
        GoldFishEncode = GoldFishEncode & Hex(n)
        ibin = ybin
    Next
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function

'This Function is for converting Decimal to Binary
Private Function Dec2Bin(dec As Integer) As String
Const mstrMethod As String = "Dec2Bin"
Dim Bits As String
On Error GoTo CheckErr
    Bits = ""
    Do Until dec = 0
        Bits = dec Mod 2 & Bits
        dec = dec \ 2
    Loop
    Do Until Len(Bits) Mod 8 = 0
        Bits = "0" & Bits
    Loop
    Dec2Bin = Bits
    Exit Function
CheckErr:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, mstrMethod
    LogErrorText "Error", mstrMethod, Err.Description
    'LogErrorDB "Function", mstrModule, mstrMethod, Err.Number, Err.Description
End Function
