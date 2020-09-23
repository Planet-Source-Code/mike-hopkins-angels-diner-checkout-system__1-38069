VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Angels Diner"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Exit"
      Height          =   375
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox PASSWORD 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox USERNAME 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Logon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4095
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password  >"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Username >"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   -240
      Picture         =   "frmLogin.frx":0000
      Top             =   3120
      Width           =   4830
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your username and password "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3855
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   0
      X2              =   4390
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Logon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'############################################
'############################################
'###                                      ###
'###   ANGELS DINER CHECKOUT AND TABLES   ###
'###                                      ###
'############################################
'############################################

'############################################
'#                                          #
'#     Written By MICHAEL HOPKINS           #
'#                                          #
'############################################
'#                                          #
'#     frmLogin:                            #
'#                                          #
'############################################

Private Sub cmdLogin_Click()

Dim usrCheck As Boolean, pswCheck As Boolean
Dim usr As String, psw As String
Dim usrLocation As Integer, pswLocation As Integer

'Button checks to see if the username and password are present
'in each array and then compares their locations.

'If the locations of the username and password match then the
'user is logged on.

'Returns string in text box as the username in upper-case
usr = USERNAME
usr = UCase(usr)

'Returns string in text box as the password in upper-case
psw = PASSWORD
psw = UCase(psw)

'Sets Initial value of usrCheck
usrCheck = False

'For loop searches for presence of username in array and notes
'it's location and notes one has been found
usrCtr = 0
For usrCtr = 0 To 9
    If usr = usrStore(usrCtr) Then
        usrCheck = True
        usrLocation = usrCtr
        Exit For
    End If
Next usrCtr

'Sets Initial value of pswCheck
pswCheck = False

'For loop searches for presence of password in array and notes
'it's location and notes one has been found
For pswCtr = 0 To 9
    If psw = pswStore(pswCtr) Then
        pswCheck = True
        pswLocation = pswCtr
        Exit For
    End If
Next pswCtr
    
    
'All below IF statements will determine whether user entered a
'matching username and password, and will take course of action
'according to usrlocation, pswlocation, usrcheck and pswCheck.

If (usrLocation = pswLocation) = False Then
    MsgBox "Invalid Username or Password", vbInformation, "Try Again!"
    USERNAME = ""
    PASSWORD = ""
    USERNAME.SetFocus
    Exit Sub
End If

If (usrCheck = False) Or (pswCheck = False) Then
    MsgBox "Invalid Username or Password", vbInformation, "Try Again!"
    USERNAME = ""
    PASSWORD = ""
    USERNAME.SetFocus
    Exit Sub
End If

If (USERNAME = "") And (PASSWORD = "") Then
    Exit Sub
End If

If (usrCheck = True) And (pswCheck = True) Then
    If usrLocation = pswLocation Then
        frmMain.Show
        frmLogin.Hide
    End If
End If

End Sub

Private Sub cmdAbout_Click()

MsgBox "This program is written by Michael Hopkins" _
        , vbInformation, "Angels Diner"

End Sub

Private Sub cmdExit_Click()

Deploy.quit

End Sub

Private Sub Form_Load()

'Load the user information at start-up
Deploy.usrLoad

End Sub

