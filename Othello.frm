VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Othello 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello"
   ClientHeight    =   5985
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6555
   Icon            =   "Othello.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   63
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   74
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   62
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   73
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   61
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   72
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   60
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   71
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   59
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   70
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   58
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   69
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   57
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   68
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   56
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   67
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   55
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   66
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   54
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   65
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   53
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   64
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   52
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   63
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   51
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   62
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   50
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   61
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   49
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   60
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   48
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   59
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   47
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   58
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   46
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   57
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   45
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   56
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   44
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   55
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   43
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   54
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   42
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   53
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   41
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   52
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   40
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   51
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   39
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   50
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   38
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   49
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   37
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   48
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   36
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   47
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   35
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   46
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   34
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   45
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   33
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   44
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   32
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   43
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   31
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   42
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   30
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   41
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   29
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   40
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   28
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   39
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   27
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   38
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   26
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   25
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   36
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   24
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   35
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   23
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   34
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   22
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   21
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   32
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   20
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   31
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   19
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   30
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   18
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   29
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   17
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   28
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   16
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   27
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   15
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   26
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   14
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   25
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   13
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   24
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   12
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   23
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   11
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   10
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   20
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   8
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   19
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   7
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   3000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox Othello 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox white 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4080
      Picture         =   "Othello.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox black 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4080
      Picture         =   "Othello.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   3000
      Width           =   495
   End
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5730
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PbProg 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ListBox Lstprog 
      Height          =   1230
      ItemData        =   "Othello.frx":0CC6
      Left            =   120
      List            =   "Othello.frx":0CC8
      TabIndex        =   4
      Top             =   4080
      Width           =   6375
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Created By Kevin Pfister"
      Height          =   255
      Left            =   4080
      TabIndex        =   75
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblcom 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblplayer 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lbltwo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblone 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "OTHELLO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   2385
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnunew 
         Caption         =   "New Game"
         Begin VB.Menu mnX 
            Caption         =   "Player as X"
         End
         Begin VB.Menu mnuO 
            Caption         =   "Player as O"
         End
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Othello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Othello Created By Kevin Pfister
'This is a totally orginal version of Othello It includes excellent AI, which so far
'i have not been able to win against, please vote or leave a comment about this or if
'there is a bug which i should know about.

'I know this is not the best coding espescially the number of nested loops, but i only took
'1 hour making this. I hope you enjoy playing this.


Dim ox(1 To 8, 1 To 8), hold(1 To 8, 1 To 8), checking(1 To 8, 1 To 8), playas, goodshot

Private Sub CmdExit_Click()
    End 'exit the program
End Sub

Private Sub Form_Load()
    playas = 1  'Play as Black
    startup 'Call the Startup sub
End Sub

Private Sub mnuExit_Click()
    End 'Exit the program
End Sub

Private Sub mnuO_Click()
    playas = 2  'Play as white
    startup 'Call the Startup sub
End Sub

Private Sub mnX_Click()
    playas = 1  'Play as Black
    startup 'Call the Startup sub
End Sub

Private Sub Othello_click(Index As Integer)
    a = Int(Index / 8) + 1
    b = Index - (a * 8) + 9
    If ox(a, b) = 0 Then
        Call place(a, b, playas)
        If goodshot = 1 Then
            comp (3 - playas)
        Else
            MsgBox ("Invalid Move")
        End If
    Else
        MsgBox ("Invalid Move")
    End If
    For c = 1 To 8
        For d = 1 To 8
            If ox(c, d) = 1 Then
                ones = ones + 1
            ElseIf ox(c, d) = 2 Then
                twos = twos + 1
            End If
        Next d
    Next c
    lblone.Caption = "  " + Str$(ones)  'change the score box
    lbltwo.Caption = "  " + Str$(twos)  'change the score box
    If ones + twos = 64 Then
        Lstprog.AddItem ("Game Over")
        If ones > twos Then
            Lstprog.AddItem ("Computer Wins")
            MsgBox ("Computer Wins")
        ElseIf ones < twos Then
            Lstprog.AddItem ("Player Wins")
            MsgBox ("Player Wins")
        ElseIf ones = twos Then
            Lstprog.AddItem ("Draw")
            MsgBox ("Draw")
        End If
    End If
End Sub

Sub startup()
    For c = 1 To 8
        For d = 1 To 8
            ox(c, d) = 0
        Next
    Next
    lblone.Caption = ""
    lbltwo.Caption = ""
    If playas = 1 Then
        lblplayer.Caption = "Player = Black"
        lblcom.Caption = "Computer = White"
    Else
        lblplayer.Caption = "Player = White"
        lblcom.Caption = "Computer = Black"
    End If
    Lstprog.Clear
    ox(4, 4) = 1
    ox(5, 4) = 2
    ox(4, 5) = 1
    ox(5, 5) = 2
    For c = 1 To 8
        For d = 1 To 8
            If ox(c, d) = 0 Then
                Othello(((c * 8) - 8) + d - 1).Picture = LoadPicture("")
            ElseIf ox(c, d) = 1 Then
                Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
            ElseIf ox(c, d) = 2 Then
                Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
            End If
        Next
    Next
    stat.Panels(1) = "Status : Player's Turn"
End Sub

Sub comp(go)
    stat.Panels(1) = "Status : Computer"
    Lstprog.AddItem ("Computer's Turn")
    PbProg.Max = 8
    For a = 1 To 8
        PbProg = a
        For b = 1 To 8
            hold(a, b) = 0
        Next
    Next
    For a = 1 To 8
        PbProg = a
        For b = 1 To 8
            If ox(a, b) = 0 Then
                Call check(a, b, go)
            End If
        Next
    Next
    winnera = 1
    winnerb = 1
    wintot = 0
    stat.Panels(1) = "Status : Thinking..."
    For a = 1 To 8
        For b = 1 To 8
            If a = 1 And b = 1 And hold(a, b) > 0 Then
                winnera = a
                winnerb = b
                wintot = 999
                Lstprog.AddItem ("Thinking...")
            ElseIf a = 1 And b = 8 And hold(a, b) > 0 Then
                winnera = a
                winnerb = b
                wintot = 999
                Lstprog.AddItem ("Thinking...")
            ElseIf a = 8 And b = 1 And hold(a, b) > 0 Then
                winnera = a
                winnerb = b
                wintot = 999
                Lstprog.AddItem ("Thinking...")
            ElseIf a = 8 And b = 8 And hold(a, b) > 0 Then
                winnera = a
                winnerb = b
                wintot = 999
                Lstprog.AddItem ("Thinking...")
            ElseIf hold(a, b) > wintot Then
                winnera = a
                winnerb = b
                wintot = hold(a, b)
                Lstprog.AddItem ("Thinking...")
            End If
        Next
    Next
    If winnera = 1 And winnerb = 1 And wintot = 0 Then
        Beep
        Lstprog.AddItem ("Unable to find Move")
    Else
        stat.Panels(1) = "Status : Playing Move"
        Lstprog.AddItem ("Chosen Move " + Str$(winnera) + "," + Str$(winnerb))
        Call place(winnera, winnerb, go)
    End If
    stat.Panels(1) = "Status : Player's Turn"
End Sub

Sub check(a, b, go)
    total = 0
    If a - 1 > 0 And b - 1 > 0 Then
        If ox(a - 1, b - 1) = go Then
            crossed = 1
            For found = 1 To 8
                If a - (1 + found) > 0 And b - (1 + found) > 0 Then
                    If ox(a - (1 + found), b - (1 + found)) = go Then
                        crossed = crossed + 1
                    ElseIf ox(a - (1 + found), b - (1 + found)) = 3 - go Then
                        total = total + crossed
                    ElseIf ox(a - (1 + found), b - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If a > 0 And b - 1 > 0 Then
        If ox(a, b - 1) = go Then
            crossed = 1
            For found = 1 To 8
                If a > 0 And b - (1 + found) > 0 Then
                    If ox(a, b - (1 + found)) = go Then
                        crossed = crossed + 1
                    ElseIf ox(a, b - (1 + found)) = 3 - go Then
                        total = total + crossed
                    ElseIf ox(a, b - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If

    If a + 1 < 8 And b - 1 > 0 Then
        If ox(a + 1, b - 1) = go Then
            crossed = 1
            For found = 1 To 8
                If a + (1 + found) < 8 And b - (1 + found) > 0 Then
                    If ox(a + (1 + found), b - (1 + found)) = go Then
                        crossed = crossed + 1
                    ElseIf ox(a + (1 + found), b - (1 + found)) = 3 - go Then
                        total = total + crossed
                    ElseIf ox(a + (1 + found), b - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If

    If a - 1 > 0 And b > 0 Then
        If ox(a - 1, b) = go Then
            crossed = 1
            For found = 1 To 8
                If a - (1 + found) > 0 And b > 0 Then
                    If ox(a - (1 + found), b) = go Then
                        crossed = crossed + 1
                    ElseIf ox(a - (1 + found), b) = 3 - go Then
                        total = total + crossed
                    ElseIf ox(a - (1 + found), b) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If a + 1 < 8 And b > 0 Then
        If ox(a + 1, b) = go Then
            crossed = 1
            For found = 1 To 8
                If a + (1 + found) < 8 And b > 0 Then
                    If ox(a + (1 + found), b) = go Then
                        crossed = crossed + 1
                    ElseIf ox(a + (1 + found), b) = 3 - go Then
                        total = total + crossed
                    ElseIf ox(a + (1 + found), b) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If a - 1 > 0 And b + 1 < 8 Then
        If ox(a - 1, b + 1) = go Then
            crossed = 1
            For found = 1 To 8
                If a - (1 + found) > 0 And b + (1 + found) < 8 Then
                    If ox(a - (1 + found), b + (1 + found)) = go Then
                        crossed = crossed + 1
                    ElseIf ox(a - (1 + found), b + (1 + found)) = 3 - go Then
                        total = total + crossed
                    ElseIf ox(a - (1 + found), b + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    
    If a > 0 And b + 1 < 8 Then
        If ox(a, b + 1) = go Then
            crossed = 1
            For found = 1 To 8
                If a > 0 And b + (1 + found) < 8 Then
                    If ox(a, b + (1 + found)) = go Then
                        crossed = crossed + 1
                    ElseIf ox(a, b + (1 + found)) = 3 - go Then
                        total = total + crossed
                    ElseIf ox(a, b + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If

    If a + 1 < 8 And b + 1 < 8 Then
        If ox(a + 1, b + 1) = go Then
            crossed = 1
            For found = 1 To 8
                If a + (1 + found) < 8 And b + (1 + found) < 8 Then
                    If ox(a + (1 + found), b + (1 + found)) = go Then
                        crossed = crossed + 1
                    ElseIf ox(a + (1 + found), b + (1 + found)) = 3 - go Then
                        total = total + crossed
                    ElseIf ox(a + (1 + found), b + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    hold(a, b) = total
End Sub

Sub place(a, b, go)
    win = 0
    If a - 1 > 0 And b - 1 > 0 Then
        If ox(a - 1, b - 1) = go Then
            checking(a - 1, b - 1) = 1
            For found = 1 To 8
                If a - (1 + found) > 0 And b - (1 + found) > 0 Then
                    If ox(a - (1 + found), b - (1 + found)) = go Then
                        checking(a - (1 + found), b - (1 + found)) = 1
                    ElseIf ox(a - (1 + found), b - (1 + found)) = 3 - go Then
                        win = 1
                        For c = 1 To 8
                            For d = 1 To 8
                                If checking(c, d) = 1 Then
                                    If go = 1 Then
                                        Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
                                        ox(c, d) = 2
                                    Else
                                        Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
                                        ox(c, d) = 1
                                    End If
                                End If
                            Next
                        Next
                    ElseIf ox(a - (1 + found), b - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
     For c = 1 To 8
        For d = 1 To 8
            If checking(c, d) = 1 Then
                checking(c, d) = 0
            End If
        Next
    Next
    If a > 0 And b - 1 > 0 Then
        If ox(a, b - 1) = go Then
            checking(a, b - 1) = 1
            For found = 1 To 8
                If a > 0 And b - (1 + found) > 0 Then
                    If ox(a, b - (1 + found)) = go Then
                        checking(a, b - (1 + found)) = 1
                    ElseIf ox(a, b - (1 + found)) = 3 - go Then
                        win = 1
                        For c = 1 To 8
                            For d = 1 To 8
                                If checking(c, d) = 1 Then
                                    If go = 1 Then
                                        Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
                                        ox(c, d) = 2
                                    Else
                                        Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
                                        ox(c, d) = 1
                                    End If
                                End If
                            Next
                        Next
                    ElseIf ox(a, b - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    For c = 1 To 8
        For d = 1 To 8
            If checking(c, d) = 1 Then
                checking(c, d) = 0
            End If
        Next
    Next
    If a + 1 < 8 And b - 1 > 0 Then
        If ox(a + 1, b - 1) = go Then
            checking(a + 1, b - 1) = 1
            For found = 1 To 8
                If a + (1 + found) < 8 And b - (1 + found) > 0 Then
                    If ox(a + (1 + found), b - (1 + found)) = go Then
                          checking(a + (1 + found), b - (1 + found)) = 1
                    ElseIf ox(a + (1 + found), b - (1 + found)) = 3 - go Then
                        win = 1
                        For c = 1 To 8
                            For d = 1 To 8
                                If checking(c, d) = 1 Then
                                    If go = 1 Then
                                        Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
                                        ox(c, d) = 2
                                    Else
                                        Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
                                        ox(c, d) = 1
                                    End If
                                End If
                            Next
                        Next
                    ElseIf ox(a + (1 + found), b - (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    For c = 1 To 8
        For d = 1 To 8
            If checking(c, d) = 1 Then
                checking(c, d) = 0
            End If
        Next
    Next
    If a - 1 > 0 And b > 0 Then
        If ox(a - 1, b) = go Then
            checking(a - 1, b) = 1
            For found = 1 To 8
                If a - (1 + found) > 0 And b > 0 Then
                    If ox(a - (1 + found), b) = go Then
                        checking(a - (1 + found), b) = 1
                    ElseIf ox(a - (1 + found), b) = 3 - go Then
                        win = 1
                        For c = 1 To 8
                            For d = 1 To 8
                                If checking(c, d) = 1 Then
                                    If go = 1 Then
                                        Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
                                        ox(c, d) = 2
                                    Else
                                        Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
                                        ox(c, d) = 1
                                    End If
                                End If
                            Next
                        Next
                    ElseIf ox(a - (1 + found), b) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    For c = 1 To 8
        For d = 1 To 8
            If checking(c, d) = 1 Then
                checking(c, d) = 0
            End If
        Next
    Next
    If a + 1 < 8 And b > 0 Then
        If ox(a + 1, b) = go Then
            checking(a + 1, b) = 1
            For found = 1 To 8
                If a + (1 + found) < 8 And b > 0 Then
                    If ox(a + (1 + found), b) = go Then
                        checking(a + (1 + found), b) = 1
                    ElseIf ox(a + (1 + found), b) = 3 - go Then
                        win = 1
                        For c = 1 To 8
                            For d = 1 To 8
                                If checking(c, d) = 1 Then
                                    If go = 1 Then
                                        Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
                                        ox(c, d) = 2
                                    Else
                                        Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
                                        ox(c, d) = 1
                                    End If
                                End If
                            Next
                        Next
                    ElseIf ox(a + (1 + found), b) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    For c = 1 To 8
        For d = 1 To 8
            If checking(c, d) = 1 Then
                checking(c, d) = 0
            End If
        Next
    Next
    If a - 1 > 0 And b + 1 < 8 Then
        If ox(a - 1, b + 1) = go Then
            checking(a - 1, b + 1) = 1
            For found = 1 To 8
                If a - (1 + found) > 0 And b + (1 + found) < 8 Then
                    If ox(a - (1 + found), b + (1 + found)) = go Then
                        checking(a - (1 + found), b + (1 + found)) = 1
                    ElseIf ox(a - (1 + found), b + (1 + found)) = 3 - go Then
                        win = 1
                        For c = 1 To 8
                            For d = 1 To 8
                                If checking(c, d) = 1 Then
                                    If go = 1 Then
                                        Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
                                        ox(c, d) = 2
                                    Else
                                        Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
                                        ox(c, d) = 1
                                    End If
                                End If
                            Next
                        Next
                    ElseIf ox(a - (1 + found), b + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    For c = 1 To 8
        For d = 1 To 8
            If checking(c, d) = 1 Then
                checking(c, d) = 0
            End If
        Next
    Next
    If a > 0 And b + 1 < 8 Then
        If ox(a, b + 1) = go Then
            checking(a, b + 1) = 1
            For found = 1 To 8
                If a > 0 And b + (1 + found) < 8 Then
                    If ox(a, b + (1 + found)) = go Then
                        checking(a, b + (1 + found)) = 1
                    ElseIf ox(a, b + (1 + found)) = 3 - go Then
                        win = 1
                        For c = 1 To 8
                            For d = 1 To 8
                                If checking(c, d) = 1 Then
                                    If go = 1 Then
                                        Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
                                        ox(c, d) = 2
                                    Else
                                        Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
                                        ox(c, d) = 1
                                    End If
                                End If
                            Next
                        Next
                    ElseIf ox(a, b + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    For c = 1 To 8
        For d = 1 To 8
            If checking(c, d) = 1 Then
                checking(c, d) = 0
            End If
        Next
    Next
    If a + 1 < 8 And b + 1 < 8 Then
        If ox(a + 1, b + 1) = go Then
            checking(a + 1, b + 1) = 1
            For found = 1 To 8
                If a + (1 + found) < 8 And b + (1 + found) < 8 Then
                    If ox(a + (1 + found), b + (1 + found)) = go Then
                        checking(a + (1 + found), b + (1 + found)) = 1
                    ElseIf ox(a + (1 + found), b + (1 + found)) = 3 - go Then
                        win = 1
                        For c = 1 To 8
                            For d = 1 To 8
                                If checking(c, d) = 1 Then
                                    If go = 1 Then
                                        Othello(((c * 8) - 8) + d - 1).Picture = black.Picture
                                        ox(c, d) = 2
                                    Else
                                        Othello(((c * 8) - 8) + d - 1).Picture = white.Picture
                                        ox(c, d) = 1
                                    End If
                                End If
                            Next
                        Next
                    ElseIf ox(a + (1 + found), b + (1 + found)) = 0 Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End If
    For c = 1 To 8
        For d = 1 To 8
            If checking(c, d) = 1 Then
                checking(c, d) = 0
            End If
        Next
    Next
    If win = 1 Then
        checking(a, b) = 1
        If go = 1 Then
            Othello(((a * 8) - 8) + b - 1).Picture = black.Picture
            ox(a, b) = 2
        Else
            Othello(((a * 8) - 8) + b - 1).Picture = white.Picture
            ox(a, b) = 1
        End If
        goodshot = 1
    ElseIf win = 0 Then
        goodshot = 0
    End If
End Sub

