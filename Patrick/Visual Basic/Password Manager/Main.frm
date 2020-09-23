VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Manager"
   ClientHeight    =   2040
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox MyIcn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1680
      Picture         =   "Main.frx":1042
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame Stored 
      Caption         =   "stored accounts"
      Height          =   1845
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4215
      Begin VB.CommandButton Add 
         Caption         =   "add account"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   1485
         Width           =   1215
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "options"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1485
         Width           =   855
      End
      Begin VB.TextBox str 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox str 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox str 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   5
         Top             =   720
         Width           =   3705
      End
      Begin VB.ComboBox cStor 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblPASSWORD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "password:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label lblUSERNAME 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "username:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   765
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "url:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   4080
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   4080
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Menu mOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mPWProtect 
         Caption         =   "Password Protect"
      End
      Begin VB.Menu mPassword 
         Caption         =   "Password.."
      End
      Begin VB.Menu mnuDash0 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Add_Click()
Dim Title As String
Title = InputBox("Enter a description for this account:", "Account Description")
If Title = "" Then Exit Sub
NumAccounts = NumAccounts + 1
Account(NumAccounts).Title = Title
cMg.AddItem Title
cMg.ListIndex = cMg.ListCount - 1
End Sub

Private Sub cmdOptions_Click()
PopupMenu mOptions, , cmdOptions.Left + Stored.Left, Stored.Top + cmdOptions.Top + cmdOptions.Height
End Sub

Private Sub cStor_Click()
Dim X As Integer
If cStor.ListIndex < 0 Then Exit Sub
X = cStor.ListIndex + 1
str(0).text = Account(X).URL
str(1).text = Account(X).Username
str(2).text = Account(X).Password
End Sub

Private Sub cStor_Scroll()
cStor_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim message As Long, PW As String
On Error Resume Next
message = X / Screen.TwipsPerPixelX
If message = &H203 Then
    If PWProtect = True Then
        PW = InputBox("Please enter your password:", "Password Required")
        If LCase(PW) <> LCase(Password) Then Exit Sub
    End If
    Me.WindowState = vbNormal
    RemoveFromTray
    Me.Visible = True
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
    AddToTray Me, "Password Manager", MyIcn
    Me.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Save
SaveAccounts
End Sub

Private Sub mExit_Click()
Unload Me
End
End Sub

Private Sub mPassword_Click()
Dim PW As String
PW = InputBox("Enter phrase to password-protect Password Manager:", "Password Protection", Password)
If PW = "" Then Exit Sub
Password = PW
End Sub

Private Sub mPWProtect_Click()
mPWProtect.Checked = Not mPWProtect.Checked
PWProtect = mPWProtect.Checked
End Sub

Private Sub str_Change(Index As Integer)
Select Case Index
    Case 0
        UpdateAccount cStor.ListIndex, "", str(0).text, "", ""
    Case 1
        UpdateAccount cStor.ListIndex, "", "", str(1).text, ""
    Case 2
        UpdateAccount cStor.ListIndex, "", "", "", str(2).text
End Select
End Sub
