VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Angels Diner"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox optCash 
      BackColor       =   &H80000009&
      Caption         =   "Cash"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   67
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End Purchases"
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
      Left            =   3360
      TabIndex        =   58
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdTable 
      Caption         =   "View TAS"
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
      Left            =   7320
      TabIndex        =   57
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Receipt"
      Enabled         =   0   'False
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
      Left            =   4680
      TabIndex        =   56
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Item"
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
      Left            =   6000
      TabIndex        =   55
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   6120
      TabIndex        =   53
      Top             =   1320
      Width           =   2535
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ListBox lstFood 
         Height          =   2460
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         TabIndex        =   54
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "val 4"
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
      Index           =   3
      Left            =   5280
      TabIndex        =   48
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "val 3"
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
      Index           =   2
      Left            =   4560
      TabIndex        =   47
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "val 2"
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
      Index           =   1
      Left            =   3840
      TabIndex        =   46
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdValue 
      Caption         =   "val 1"
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
      Index           =   0
      Left            =   3120
      TabIndex        =   45
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "cBurger"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   44
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Fries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   43
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "1/2 lb"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   42
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "1/4 lb"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   41
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Water"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      HelpContextID   =   91
      Index           =   18
      Left            =   1680
      TabIndex        =   40
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Worthie"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   960
      TabIndex        =   39
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "sBow"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   240
      TabIndex        =   38
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Soft"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   240
      TabIndex        =   37
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Tikka"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   36
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "oRings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   35
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Chik C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2400
      TabIndex        =   34
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Shakes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   960
      TabIndex        =   33
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Bud"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1680
      TabIndex        =   32
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Tea"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1680
      TabIndex        =   31
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "lasagne"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   960
      TabIndex        =   30
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Ceaser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   29
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Wedge"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   28
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Coffee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2400
      TabIndex        =   27
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Carling"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   2400
      TabIndex        =   26
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtLabel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdSendReceipt 
      Caption         =   "Enter"
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
      Left            =   5040
      TabIndex        =   23
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
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
      Left            =   3600
      TabIndex        =   22
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
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
      Left            =   4080
      TabIndex        =   21
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
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
      Left            =   4560
      TabIndex        =   20
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "*"
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
      Left            =   4560
      TabIndex        =   19
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdSubtract 
      Caption         =   "-"
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
      Left            =   4560
      TabIndex        =   18
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelUser 
      Caption         =   "C"
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
      Left            =   3840
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "0"
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
      Index           =   0
      Left            =   3120
      MaskColor       =   &H00FF0000&
      TabIndex        =   16
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdAddition 
      Caption         =   "+"
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
      Left            =   4560
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "9"
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
      Index           =   9
      Left            =   4080
      MaskColor       =   &H00FF0000&
      TabIndex        =   14
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "8"
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
      Index           =   8
      Left            =   3600
      MaskColor       =   &H00FF0000&
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "7"
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
      Index           =   7
      Left            =   3120
      MaskColor       =   &H00FF0000&
      TabIndex        =   12
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "6"
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
      Index           =   6
      Left            =   4080
      MaskColor       =   &H00FF0000&
      TabIndex        =   11
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "5"
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
      Index           =   5
      Left            =   3600
      MaskColor       =   &H00FF0000&
      TabIndex        =   10
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "4"
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
      Index           =   4
      Left            =   3120
      MaskColor       =   &H00FF0000&
      TabIndex        =   9
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "3"
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
      Index           =   3
      Left            =   4080
      MaskColor       =   &H00FF0000&
      TabIndex        =   8
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "2"
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
      Index           =   2
      Left            =   3600
      MaskColor       =   &H00FF0000&
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "1"
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
      Index           =   1
      Left            =   3120
      MaskColor       =   &H00FF0000&
      TabIndex        =   6
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   405
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdPassword 
      Caption         =   "Change Password"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   840
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "Add New User"
      Enabled         =   0   'False
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
      Left            =   1440
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
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
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Checkout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   52
      Top             =   1320
      Width           =   5895
      Begin VB.CommandButton cmdRoot 
         Caption         =   "sqr"
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
         Left            =   5400
         TabIndex        =   64
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdPercent 
         Caption         =   "%"
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
         Left            =   4920
         TabIndex        =   63
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "CE"
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
         TabIndex        =   62
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdUserExit 
         Caption         =   "Cancel"
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
         Left            =   4920
         TabIndex        =   61
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdUserDef 
         Caption         =   "UserDef"
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
         Left            =   4920
         TabIndex        =   60
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   65
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000009&
      Height          =   1335
      Left            =   3240
      TabIndex        =   66
      Top             =   4680
      Width           =   5415
      Begin VB.CheckBox optCheque 
         BackColor       =   &H80000009&
         Caption         =   "Cheque"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   69
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox optCreditCard 
         BackColor       =   &H80000009&
         Caption         =   "Credit Card"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   68
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   2265
      Left            =   -600
      Picture         =   "frmMain.frx":0004
      Top             =   5760
      Width           =   9450
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   9000
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Angels Diner Checkout"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   51
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblDate 
      BackColor       =   &H80000009&
      Caption         =   "Label3"
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
      Left            =   7320
      TabIndex        =   50
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblTime 
      BackColor       =   &H80000009&
      Caption         =   "Label2"
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
      Left            =   6240
      TabIndex        =   49
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Welcome"
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
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
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
'#     frmMain:                             #
'#                                          #
'############################################

'Welcomes user
Dim strName As String
'Store for calculator
Dim store As Double
'Decimal Places
Dim DecimalPositions As Integer
'Checks for decimal
Dim checkDecimal As Boolean
'Credit Card # (for checking)
Dim strCCNumber As String
'Operator wanted
Dim operStore As String

Private Sub cmdAddition_Click()
    
'Sub for Addition function of calculator

'Stores current value in memory
store = txtValue.Text

'Reset display
txtValue.Text = "0."

'Sets function wanted
operStore = "add"

'Resets decimal functions
checkDecimal = False
DecimalPositions = 1

End Sub

Private Sub cmdCancel_Click()

'Sub cancels last item entered

'If there is no item present on display then exit
If txtLabel = "" Then
    Exit Sub
End If

'Reset displays
txtLabel.Text = ""
txtValue.Text = "0."

'Items in list
Dim intItems As Integer

intItems = lstFood.ListCount

'If items are present then last is removed
If intItems = 0 Then
    Exit Sub
    Else
    lstFood.RemoveItem (intItems - 1)
End If

'Recalculate price
Deploy.totPrice

End Sub

Private Sub CmdCancelUser_Click()

'Clears all memory in User Defined calculations
Deploy.cancelAll
checkDecimal = False
DecimalPositions = 1

End Sub

Private Sub cmdClear_Click()

'Clears active number in User Defined calculations
txtValue.Text = "0.00"

End Sub

Private Sub cmdDecimal_Click()

'Enables decimal calculations
checkDecimal = True

End Sub

Private Sub cmdDivide_Click()

'Sub for Division function of calculator

'Stores current value in memory
store = txtValue.Text

'Reset display
txtValue.Text = "0."

'Sets function wanted
operStore = "divide"

'Resets decimal functions
checkDecimal = False
DecimalPositions = 1

End Sub

Private Sub cmdEquals_Click()

'Sub calculates new value by considering
'operator and value stored in memory

'Exits sub if no operator is present
If operStore = "" Then
    Exit Sub
End If

'Sets function in operStore to 'add'
'then adds value in memory to value in display
If operStore = "add" Then
    txtValue.Text = (Val(store) + Val(txtValue.Text))
End If

'Sets function in operStore to 'subtract'
'then subtracts value in memory from value in display
If operStore = "subtract" Then
    txtValue.Text = (Val(store) - Val(txtValue.Text))
End If

'Sets function in operStore to 'multiply'
'then multiplies value in memory with value in display
If operStore = "multiply" Then
    txtValue.Text = (Val(store) * Val(txtValue.Text))
End If

'Sets function in operStore to 'divide'
'then divides value in memory into value in display
If operStore = "divide" Then
    If txtValue.Text = "0.00" Then
        txtValue.Text = "Error"
        Else
        txtValue.Text = (Val(store) / Val(txtValue.Text))
    End If
End If

'Next section of the sub determines whether
'answer in an integer or decimal-based number
'and adds a decimal place accordingly.

Dim strLenth As Integer
Dim intCtr As Integer
Dim dec As Boolean

strLength = Len(txtValue.Text)

intCtr = 0
For intCtr = 1 To (strLength)
    If Mid(txtValue.Text, intCtr, 1) = "." Then
        dec = False
        Exit For
        Else
        dec = True
    End If
Next intCtr

If dec = True Then
    txtValue.Text = txtValue.Text & "."
    Else
    txtValue.Text = Val(txtValue.Text)
End If

'Stores new answer in memory
store = Val(txtValue.Text)

'resets the calculator for another calculation
operStore = ""
checkDecimal = False
DecimalPositions = 0

End Sub

Private Sub cmdExit_Click()

'Unloads everything from memory
Deploy.quit

End Sub

Private Sub cmdInteger_Click(Index As Integer)

'Sub enters numbers from control array into display
Dim Ans As Double

'When checkDecimal is boolean false the
'number on the button is added to end of display
If checkDecimal = False Then
    If txtValue.Text = "" Or "0." Then
        Ans = Val(cmdInteger(Index).Caption)
        Else
        Ans = (txtValue.Text * 10) + Val(cmdInteger(Index).Caption)
    End If
End If
    
'When checkDecimal is boolean true the
'number is added to the end of the display
'and DecimalPositions increases by one
If checkDecimal = True Then
    Ans = txtValue.Text + (Val(cmdInteger(Index).Caption) / 10 ^ DecimalPositions)
    txtValue.Text = strAns
    DecimalPositions = DecimalPositions + 1
End If

txtValue.Text = Ans

End Sub

Private Sub cmdLogOut_Click()

'Logs user out of the system

Dim strYes As String

'User is asks to confirm the logout action
strYes = MsgBox("Are you sure you want to logout?", vbYesNo, "Logout")

'If the user clicks yes they are logged out, if not then sub exits
If strYes = vbYes Then
    Load frmLogin
    frmLogin.Show
    Unload frmMain
    Else
    Exit Sub
End If

End Sub

Private Sub cmdMenu_Click(Index As Integer)

'Sub for Menu control array
'enters information from array into text boxes
txtLabel.Text = menuStore(Index)
txtValue.Text = priceStore(Index)

Dim strList As String

'Adds item into list in format "NAME - PRICE"
strList = txtLabel.Text & " - " & txtValue.Text
lstFood.AddItem strList

'Keeps a numeric value in txtTotal
If txtTotal.Text = "" Then
    txtTotal.Text = "0"
End If

'Updates the total price
txtTotal.Text = txtTotal.Text + Val(txtValue.Text)

End Sub

Private Sub cmdMultiply_Click()

'Sub for Multiplication function of calculator

'Stores current value in memory
store = txtValue.Text

'Reset display
txtValue.Text = "0."

'Sets function wanted
operStore = "multiply"

'Resets decimal functions
checkDecimal = False
DecimalPositions = 1

End Sub

Private Sub cmdPassword_Click()

Dim intCtr As Integer
Dim strNewPsw As String
Dim strNewPswConf As String
Dim strOldPsw As String

Deploy.usrLoad

usr = UCase(strName)

intCtr = 0
For intCtr = 0 To 20
    If usrStore(intCtr) = usr Then
        usrLocation = intCtr
        Exit For
    End If
Next intCtr

strOldPsw = InputBox("Please Enter Your Current Password", "Password")

If strOldPsw = "" Then
    MsgBox "Password Was Not Changed", vbInformation, "Action Cancelled"
    Exit Sub
End If

strOldPsw = UCase(strOldPsw)
Do Until (UCase(strOldPsw) = pswStore(usrLocation))
    strOldPsw = InputBox("Incorrect Password, Please Try Again", "Password")
Loop

strNewPsw = InputBox("Please Enter New Password", "Password")
strNewPswConf = InputBox("Please Confirm this Password", "Password")

Do Until (strNewPsw = strNewPswConf)
    MsgBox "Passwords did not match", vbExclamation, "Please re-enter"
    strNewPsw = InputBox("Please Enter New Password Again", "Password")
    strNewPswConf = InputBox("Please Confirm this Password", "Password")
Loop

strNewPsw = UCase(strNewPsw)

pswStore(usrLocation) = strNewPsw

Open "d:\diner\users.txt" For Input As #1

Dim intmax As Integer

intmax = 0
Do Until (EOF(1) = True)
    Input #1, usrStore(), pswStore()
    intmax = intmax + 1
Loop

Close #1

Open "d:\diner\users.txt" For Output As #1
intCtr = 0
For intCtr = 0 To (intmax - 1)
    Write #1, usrStore(intCtr), pswStore(intCtr)
Next intCtr

Close #1

MsgBox "Password Succesfully Changed", vbInformation, "Password"

End Sub


Private Sub cmdPercent_Click()

txtValue.Text = txtValue.Text / 100

End Sub

Private Sub cmdPrint_Click()

cmdPrint.Enabled = False

Deploy.cancelAll
frmMain.txtLabel.Text = ""
checkDecimal = False
DecimalPositions = 1
lstFood.Clear
txtTotal.Text = "0.00"
cmdRemove.Enabled = True
cmdPrint.Enabled = False
cmdEnd.Enabled = True

optCash.Value = 0
optCreditCard.Value = 0
optCheque.Value = 0

End Sub

Private Sub cmdRemove_Click()

Dim lstCount As Integer

ListCount = lstFood.ListCount
    Do While ListCount > 0
    ListCount = ListCount - 1
    If lstFood.Selected(ListCount) = True Then
        lstFood.RemoveItem (ListCount)
    End If
Loop

Deploy.totPrice

End Sub

Private Sub cmdRoot_Click()

txtValue.Text = Sqr(txtValue.Text)

End Sub

Private Sub cmdSendReceipt_Click()

If txtValue.Text = "0.00" Then
    Exit Sub
End If

If Mid(txtValue.Text, Len(txtValue.Text) - 2, 1) = "." Then
    If txtLabel.Text = "User Defined" Then
        Dim strList As String
        strList = txtLabel.Text & " - " & txtValue.Text
        lstFood.AddItem strList
    End If
    Deploy.userDefined
    Deploy.totPrice
    Exit Sub
End If

MsgBox "This is not a valid price", vbInformation, "Try Again"
Deploy.cancelAll
frmMain.txtLabel.Text = ""
checkDecimal = False
DecimalPositions = 1
Deploy.userDefined

End Sub

Private Sub cmdSubtract_Click()

'Sub for Subtraction function of calculator

'Stores current value in memory
store = txtValue.Text

'Reset display
txtValue.Text = "0."

'Sets function wanted
operStore = "subtract"

'Resets decimal functions
checkDecimal = False
DecimalPositions = 1

End Sub

Private Sub cmdTable_Click()

MsgBox "To be implemented in CP4", vbInformation, "TAS"

End Sub

Private Sub cmdUser_Click()

'Initiates sequence to add a new user
Deploy.addUsers

End Sub

Private Sub cmdUserDef_Click()

'Sets state of calculator buttons
With frmMain
    .cmdAddition.Enabled = True
    .cmdSubtract.Enabled = True
    .cmdMultiply.Enabled = True
    .cmdDivide.Enabled = True
    .cmdSendReceipt.Enabled = True
    .cmdDecimal.Enabled = True
    .cmdClear.Enabled = True
    .cmdEquals.Enabled = True
    .cmdCancelUser.Enabled = True
    .cmdUserExit.Enabled = True
    .cmdRoot.Enabled = True
    .cmdPercent.Enabled = True
End With

Dim intCtr As Integer

intCtr = 0
For intCtr = 0 To 9
    frmMain.cmdInteger(intCtr).Enabled = True
Next intCtr
    
'Sets initial state of calculator memory
Deploy.cancelAll
checkDecimal = False
DecimalPositions = 1

'Show User Defined in display
txtLabel.Text = "User Defined"

End Sub

Private Sub cmdUserExit_Click()

'Sub cancels User Defined value
Deploy.cancelAll
frmMain.txtLabel.Text = ""
checkDecimal = False
DecimalPositions = 1
Deploy.userDefined

End Sub

Private Sub cmdEnd_Click()

'Sub initiates end of order
'If nothing has been bought then
'the sub is exited
If txtTotal.Text = "0.00" Then
    Exit Sub
End If

Deploy.cancelAll
frmMain.txtLabel.Text = ""
checkDecimal = False
DecimalPositions = 1

txtLabel.Text = "Total Price"
txtValue.Text = txtTotal.Text

'Enable payment methods
optCash.Enabled = True
optCreditCard.Enabled = True
optCheque.Enabled = True

cmdRemove.Enabled = False
cmdEnd.Enabled = False


End Sub

Private Sub Form_Load()

'Sets initial states of Command Buttons
With frmMain
    .cmdAddition.Enabled = False
    .cmdSubtract.Enabled = False
    .cmdMultiply.Enabled = False
    .cmdDivide.Enabled = False
    .cmdSendReceipt.Enabled = False
    .cmdDecimal.Enabled = False
    .cmdClear.Enabled = False
    .cmdEquals.Enabled = False
    .cmdCancelUser.Enabled = False
    .cmdUserExit.Enabled = False
    .cmdRoot.Enabled = False
    .cmdPercent.Enabled = False
End With

Dim intCtr As Integer

intCtr = 0
For intCtr = 0 To 9
    frmMain.cmdInteger(intCtr).Enabled = False
Next intCtr

'Sets state of memory
DecimalPositions = 1
checkDecimal = False
store = 0

'Sets state of text boxes
With frmMain
    .txtValue.Text = "0.00"
    .txtValue.Locked = True
    .txtLabel.Text = ""
    .txtLabel.Locked = True
    .txtTotal.Text = "0.00"
End With

'Loads name from usr in frmLogin and converts
'from format AAAAA to Aaaaa, then unloads the
'login form for memory
strName = frmLogin.USERNAME
name1 = UCase(Left(strName, 1))
name2 = LCase(Right(strName, Len(strName) - 1))
strName = name1 & name2
Label1.Caption = "Welcome " & (strName) & "!"

Unload frmLogin

'Enables add-user function if admin is logged in
If UCase(strName) = "ADMIN" Then
    cmdUser.Enabled = True
End If

'Loads time and date into form
Dim strTime As String
lblTime.Caption = Time
lblDate.Caption = Date

'Loads the menu from file into array
Deploy.menuLoad

End Sub

Private Sub optCash_Click()

'Sub for cash payment
If optCash.Value = 1 Then
    'Reset other payment buttons
    optCreditCard.Value = 0
    optCheque.Value = 0
    
    'Money given by customer
    strprice = InputBox("Cash Recieved", "Enter Amount", (frmMain.txtValue.Text))

    'Calculate change needed
    txtLabel.Text = "Change Due"
    txtValue.Text = strprice - txtValue.Text

    'Disable payment buttons
    optCash.Enabled = False
    optCreditCard.Enabled = False
    optCheque.Enabled = False
    
    'Enable buuton to allow next stage of process
    cmdPrint.Enabled = True
End If

End Sub

Private Sub optCheque_Click()

'Sub for cheque payment
If optCheque.Value = 1 Then
    'Reset other payment buttons
    optCash.Value = 0
    optCreditCard.Value = 0
    
    'Money given by customer
    strprice = InputBox("Cheque Value", "Enter Amount", (frmMain.txtValue.Text))
    
    'Calculate change needed
    txtLabel.Text = "Change Due"
    txtValue.Text = strprice - txtValue.Text

    'Disable payment buttons
    optCash.Enabled = False
    optCreditCard.Enabled = False
    optCheque.Enabled = False
    
    'Enable buuton to allow next stage of process
    cmdPrint.Enabled = True
End If

End Sub

Private Sub optCreditCard_Click()

'Sub for cheque payment
If optCreditCard.Value = 1 Then
    'Reset other payment buttons
    optCheque.Value = 0
    optCash.Value = 0
    
    'Customers Credit Card number
    strCCNumber = InputBox("Credit Card Number", "Enter Number")

    'No change from Credit Card
    txtLabel.Text = "Change Due"
    txtValue.Text = "0.00"
    
    'Disable payment buttons
    optCash.Enabled = False
    optCreditCard.Enabled = False
    optCheque.Enabled = False
    
    'Enable buuton to allow next stage of process
    cmdPrint.Enabled = True
End If

End Sub

Private Sub Timer1_Timer()

'Timer updates the time every second (interval = 1000)
strTime = Time
lblTime.Caption = strTime

End Sub

Private Sub txtTotal_Change()

'Formats the display to currency
'i.e. 34 = 34.00
txtTotal.Text = Format(txtTotal.Text, "#0.00")

End Sub

Private Sub txtValue_Change()

'Formats the display to currency
'i.e. 34 = 34.00
txtValue.Text = Format(txtValue.Text, "#0.00")

End Sub
