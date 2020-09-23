Attribute VB_Name = "pwstore"
Global PWProtect As Boolean
Global Password As String
Global Account(1 To 1000) As Acct
Global NumAccounts As Integer
Global DataFile As String
Private Type Acct
    Title As String
    URL As String
    Password As String
    Username As String
End Type

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Sub Main()
DataFile = App.Path
If Right(App.Path, 1) <> "\" Then DataFile = DataFile & "\"
DataFile = DataFile & "\data.v00"

Load
LoadAccounts
AddToTray fMain, "Password Manager", fMain.MyIcn
fMain.mPWProtect.Checked = PWProtect
End Sub

Sub LoadAccounts()
Dim Accounts As String
Dim One As String
NumAccounts = 0
Accounts = LoadFile(DataFile)
If Accounts = "" Then Exit Sub

Accounts = Crypt(Accounts)
If Right(Accounts, 2) = vbCrLf Then
    Do
        Accounts = Left(Accounts, Len(Accounts) - 2)
    Loop Until Right(Accounts, 2) <> vbCrLf
End If
If Right(Accounts, 2) <> vbCrLf Then Accounts = Accounts & vbCrLf
Do
    One = Left(Accounts, InStr(Accounts, vbCrLf) - 1)
    Accounts = Mid(Accounts, InStr(Accounts, vbCrLf) + 2, Len(Accounts))
    
    If One <> "" And One <> "çä" Then
        NumAccounts = NumAccounts + 1
        Account(NumAccounts).Title = Left(One, InStr(One, "ÿ") - 1)
        One = Mid(One, InStr(One, "ÿ") + 1, Len(One))
        
        Account(NumAccounts).URL = Left(One, InStr(One, "¯") - 1)
        One = Mid(One, InStr(One, "¯") + 1, Len(One))
        
        Account(NumAccounts).Username = Left(One, InStr(One, "õ") - 1)
        One = Mid(One, InStr(One, "õ") + 1, Len(One))
        
        Account(NumAccounts).Password = One
        fMain.cStor.AddItem Account(NumAccounts).Title
    End If
Loop Until InStr(Accounts, vbCrLf) = 0
End Sub

Sub SaveAccounts()
Dim Accounts As String
For X = 1 To NumAccounts
    Accounts = Accounts & Account(X).Title & "ÿ"
    Accounts = Accounts & Account(X).URL & "¯"
    Accounts = Accounts & Account(X).Username & "õ"
    Accounts = Accounts & Account(X).Password & vbCrLf
Next X
Accounts = Crypt(Accounts)

Dim F As Integer
F = FreeFile
Open DataFile For Output As #F
Print #F, Accounts
Close #F
End Sub

Function Crypt(text As String) As String
Dim strTempChar As String, I As Integer
For I = 1 To Len(text)
    If Asc(Mid(text, I, 1)) < 128 Then
        strTempChar = Asc(Mid$(text, I, 1)) + 128
    ElseIf Asc(Mid(text, I, 1)) > 128 Then
        strTempChar = Asc(Mid$(text, I, 1)) - 128
    End If
    Mid(text, I, 1) = Chr(strTempChar)
Next I
Crypt = text
End Function

Function LoadFile(Filename As String) As String
Dim F As Integer, Filetext As String
On Error GoTo errhandle
F = FreeFile
Open Filename For Input As #F
Filetext = Input(LOF(F), #F)
Close #F

display:
LoadFile = Filetext

Exit Function
errhandle:
Filetext = ""
Close #F
GoTo display
End Function

Sub UpdateAccount(X As Integer, Title As String, URL As String, Username As String, Password As String)
X = X + 1
If Title <> "" Then Account(X).Title = Title
If URL <> "" Then Account(X).URL = URL
If Username <> "" Then Account(X).Username = Username
If Password <> "" Then Account(X).Password = Password
End Sub




Sub INI_Write(Filename As String, Section As String, Key As String, Data As String)
WritePrivateProfileString Section, Key, Data, Filename
End Sub

Function INI_Get(Filename As String, Section As String, Key As String, DefaultData As String) As String
Dim Data As String, Tmp As Long
Data = String(750, vbNullChar)
Tmp = GetPrivateProfileString(Section, Key, DefaultData, Data, Len(Data), Filename)
INI_Get = Left(Data, InStr(Data, vbNullChar) - 1)
End Function

Sub Save()
INI_Write "C:\windows\win.ini", "PWMgr", "Default", PWProtect + 0
INI_Write "C:\windows\win.ini", "PWMgr", "Trans32", Crypt(Password)
End Sub

Sub Load()
PWProtect = INI_Get("C:\windows\win.ini", "PWMgr", "Default", 0)
Password = INI_Get("C:\windows\win.ini", "PWMgr", "Trans32", "")
If Password <> "" Then Password = Crypt(Password)
End Sub
