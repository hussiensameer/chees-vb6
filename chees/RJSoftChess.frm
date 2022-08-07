VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form RJSoftChess 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RJ SOFT CHESS"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   Icon            =   "RJSoftChess.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton PlayingOption 
      BackColor       =   &H00000000&
      Caption         =   "White"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Index           =   0
      Left            =   7830
      TabIndex        =   47
      Top             =   30
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.OptionButton PlayingOption 
      BackColor       =   &H00000000&
      Caption         =   "Black"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Index           =   1
      Left            =   8880
      TabIndex        =   46
      Top             =   30
      Width           =   1005
   End
   Begin VB.TextBox tHost 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   6240
      TabIndex        =   36
      Text            =   "192.168.0.3"
      Top             =   7320
      Width           =   1785
   End
   Begin VB.TextBox LocalIP 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   3390
      TabIndex        =   35
      Text            =   "192.168.0.3"
      Top             =   7320
      Width           =   1785
   End
   Begin VB.Timer GameTimer 
      Left            =   450
      Top             =   7290
   End
   Begin VB.TextBox txtMaxGameTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   7980
      TabIndex        =   34
      Text            =   "5"
      Top             =   3560
      Width           =   375
   End
   Begin VB.TextBox tMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1305
      Left            =   6780
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   1830
      Width           =   3045
   End
   Begin VB.TextBox tSend 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1305
      Left            =   6780
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "6"
      Top             =   4290
      Width           =   3045
   End
   Begin VB.ListBox lName 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1290
      IntegralHeight  =   0   'False
      Left            =   6750
      Sorted          =   -1  'True
      TabIndex        =   22
      Top             =   5640
      Width           =   3100
   End
   Begin MSWinsockLib.Winsock Wsck 
      Left            =   30
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.ListBox MoveList 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1290
      IntegralHeight  =   0   'False
      Left            =   6750
      TabIndex        =   24
      Top             =   5640
      Width           =   3100
   End
   Begin VB.ListBox ConnectedList 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1260
      IntegralHeight  =   0   'False
      Left            =   6780
      Sorted          =   -1  'True
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label CommandMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOARD"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Index           =   4
      Left            =   8610
      TabIndex        =   48
      Top             =   1365
      Width           =   840
   End
   Begin VB.Image CommandImage 
      Height          =   420
      Index           =   4
      Left            =   8370
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0C0C0&
      Height          =   1365
      Left            =   6750
      Top             =   4260
      Width           =   3105
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      Height          =   1365
      Left            =   6750
      Top             =   1800
      Width           =   3105
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      Height          =   405
      Left            =   7945
      Top             =   3510
      Width           =   435
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   435
      Index           =   1
      Left            =   6210
      Top             =   7290
      Width           =   1845
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   435
      Index           =   0
      Left            =   3360
      Top             =   7290
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS:"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Left            =   570
      TabIndex        =   44
      Top             =   30
      Width           =   765
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Left            =   6960
      TabIndex        =   43
      Top             =   30
      Width           =   705
   End
   Begin VB.Label ToLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Left            =   7410
      TabIndex        =   42
      Top             =   30
      Width           =   435
   End
   Begin VB.Label FromLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Left            =   6810
      TabIndex        =   41
      Top             =   30
      Width           =   435
   End
   Begin VB.Label lblMoving 
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIVING MOVE.  PLEASE STANDBY..."
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1380
      TabIndex        =   40
      Top             =   30
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.Label ConnectLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "NOT CONNECTED"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Left            =   8190
      TabIndex        =   39
      Top             =   7320
      Width           =   1635
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER:"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Left            =   5340
      TabIndex        =   38
      Top             =   7320
      Width           =   825
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOCAL:"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   345
      Left            =   2610
      TabIndex        =   37
      Top             =   7320
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   9720
      X2              =   9855
      Y1              =   3705
      Y2              =   3705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   465
      Left            =   7920
      Top             =   3480
      Width           =   1785
   End
   Begin VB.Image CursorArrow 
      Height          =   165
      Index           =   1
      Left            =   9420
      Picture         =   "RJSoftChess.frx":1BB2
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image CursorArrow 
      Height          =   165
      Index           =   0
      Left            =   9420
      Picture         =   "RJSoftChess.frx":1E04
      Top             =   3525
      Width           =   240
   End
   Begin VB.Label MaxGameTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MIN GAME"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   315
      Left            =   8415
      TabIndex        =   33
      Top             =   3555
      Width           =   990
   End
   Begin VB.Label GameTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   315
      Index           =   2
      Left            =   7440
      TabIndex        =   32
      Top             =   3930
      Width           =   495
   End
   Begin VB.Label GameTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   315
      Index           =   1
      Left            =   7440
      TabIndex        =   31
      Top             =   3180
      Width           =   495
   End
   Begin VB.Label TimerLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TIMER:"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Index           =   1
      Left            =   6750
      TabIndex        =   30
      Top             =   3930
      Width           =   645
   End
   Begin VB.Label TimerLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TIMER:"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Index           =   0
      Left            =   6750
      TabIndex        =   29
      Top             =   3180
      Width           =   645
   End
   Begin VB.Shape PreSquare 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   720
      Left            =   2670
      Top             =   8070
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LabelY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   7
      Left            =   45
      TabIndex        =   8
      Top             =   6120
      Width           =   225
   End
   Begin VB.Label LabelY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   945
      Width           =   225
   End
   Begin VB.Image OppsPiece 
      Height          =   585
      Left            =   570
      Top             =   9480
      Width           =   495
   End
   Begin VB.Label GuestBoard 
      BackStyle       =   0  'Transparent
      Caption         =   "    GUEST MODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6465
      Left            =   360
      TabIndex        =   27
      Top             =   465
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Image Master_White 
      Height          =   720
      Index           =   0
      Left            =   150
      Picture         =   "RJSoftChess.frx":2056
      Top             =   8850
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_White 
      Height          =   720
      Index           =   1
      Left            =   750
      Picture         =   "RJSoftChess.frx":3D20
      Top             =   8850
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_White 
      Height          =   720
      Index           =   2
      Left            =   1350
      Picture         =   "RJSoftChess.frx":59EA
      Top             =   8850
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_White 
      Height          =   720
      Index           =   3
      Left            =   1920
      Picture         =   "RJSoftChess.frx":76B4
      Top             =   8850
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_White 
      Height          =   720
      Index           =   4
      Left            =   2490
      Picture         =   "RJSoftChess.frx":937E
      Top             =   8850
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_White 
      Height          =   720
      Index           =   5
      Left            =   3030
      Picture         =   "RJSoftChess.frx":B048
      Top             =   8850
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_Black 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "RJSoftChess.frx":CD12
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_Black 
      Height          =   720
      Index           =   1
      Left            =   750
      Picture         =   "RJSoftChess.frx":E9DC
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_Black 
      Height          =   720
      Index           =   2
      Left            =   1350
      Picture         =   "RJSoftChess.frx":106A6
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_Black 
      Height          =   720
      Index           =   3
      Left            =   1920
      Picture         =   "RJSoftChess.frx":12370
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_Black 
      Height          =   720
      Index           =   4
      Left            =   2520
      Picture         =   "RJSoftChess.frx":1403A
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Master_Black 
      Height          =   720
      Index           =   5
      Left            =   3090
      Picture         =   "RJSoftChess.frx":15D04
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image GRG 
      Height          =   315
      Index           =   1
      Left            =   420
      Picture         =   "RJSoftChess.frx":179CE
      Top             =   10320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image GRG 
      Height          =   315
      Index           =   2
      Left            =   600
      Picture         =   "RJSoftChess.frx":17CB0
      Top             =   10320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image MouseImage 
      Height          =   480
      Index           =   0
      Left            =   780
      Picture         =   "RJSoftChess.frx":17F92
      Top             =   10320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label BoardLocation 
      Height          =   285
      Left            =   3000
      TabIndex        =   26
      Top             =   10410
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Shape FromSquare 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   720
      Left            =   300
      Top             =   8070
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label BoardNumber 
      Height          =   285
      Left            =   2310
      TabIndex        =   25
      Top             =   10410
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image Master_Blank 
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   1080
      Picture         =   "RJSoftChess.frx":18C5C
      Top             =   8040
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Shape ToSquare 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Height          =   720
      Left            =   1920
      Top             =   8070
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label LocalIPLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RJ Soft of West Tennessee "
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   6750
      TabIndex        =   21
      Top             =   7020
      Width           =   3075
   End
   Begin VB.Image WhosTurn 
      Height          =   315
      Index           =   2
      Left            =   6540
      Picture         =   "RJSoftChess.frx":1A926
      Top             =   3930
      Width           =   150
   End
   Begin VB.Image WhosTurn 
      Height          =   315
      Index           =   1
      Left            =   6540
      Picture         =   "RJSoftChess.frx":1AC08
      Top             =   3180
      Width           =   150
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   7
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":1AEEA
      Top             =   750
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   6
      Left            =   1335
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":1CBB4
      Top             =   750
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   5
      Left            =   2070
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":1E87E
      Top             =   750
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   4
      Left            =   2805
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":20548
      Top             =   750
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   3
      Left            =   3540
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":22212
      Top             =   750
      Width           =   720
   End
   Begin VB.Image Board 
      DataField       =   "C1"
      Height          =   720
      Index           =   2
      Left            =   4275
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":23EDC
      Top             =   750
      Width           =   720
   End
   Begin VB.Image Board 
      DataField       =   "B2"
      Height          =   720
      Index           =   1
      Left            =   5010
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":25BA6
      Top             =   750
      Width           =   720
   End
   Begin VB.Image Board 
      DataField       =   "A1"
      Height          =   720
      Index           =   0
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":27870
      Top             =   750
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   15
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":2953A
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   14
      Left            =   1335
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":2B204
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   13
      Left            =   2070
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":2CECE
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   12
      Left            =   2805
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":2EB98
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   11
      Left            =   3540
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":30862
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   10
      Left            =   4275
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":3252C
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   9
      Left            =   5010
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":341F6
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   8
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":35EC0
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   23
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":37B8A
      Top             =   2235
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   22
      Left            =   1335
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":39854
      Top             =   2235
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   21
      Left            =   2070
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":3B51E
      Top             =   2235
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   20
      Left            =   2805
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":3D1E8
      Top             =   2235
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   19
      Left            =   3540
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":3EEB2
      Top             =   2235
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   18
      Left            =   4275
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":40B7C
      Top             =   2235
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   17
      Left            =   5010
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":42846
      Top             =   2235
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   16
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":44510
      Top             =   2235
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   31
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":461DA
      Top             =   2970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   30
      Left            =   1335
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":47EA4
      Top             =   2970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   29
      Left            =   2070
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":49B6E
      Top             =   2970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   28
      Left            =   2805
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":4B838
      Top             =   2970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   27
      Left            =   3540
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":4D502
      Top             =   2970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   26
      Left            =   4275
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":4F1CC
      Top             =   2970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   25
      Left            =   5010
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":50E96
      Top             =   2970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   24
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":52B60
      Top             =   2970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   56
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":5482A
      Top             =   5970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   57
      Left            =   5010
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":564F4
      Top             =   5970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   58
      Left            =   4275
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":581BE
      Top             =   5970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   59
      Left            =   3540
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":59E88
      Top             =   5970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   60
      Left            =   2805
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":5BB52
      Top             =   5970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   61
      Left            =   2070
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":5D81C
      Top             =   5970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   62
      Left            =   1335
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":5F4E6
      Top             =   5970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   63
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":611B0
      Top             =   5970
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   48
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":62E7A
      Top             =   5205
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   49
      Left            =   5010
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":64B44
      Top             =   5205
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   50
      Left            =   4275
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":6680E
      Top             =   5205
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   51
      Left            =   3540
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":684D8
      Top             =   5205
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   52
      Left            =   2805
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":6A1A2
      Top             =   5205
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   53
      Left            =   2070
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":6BE6C
      Top             =   5205
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   54
      Left            =   1335
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":6DB36
      Top             =   5205
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   55
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":6F800
      Top             =   5205
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   40
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":714CA
      Top             =   4470
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   41
      Left            =   5010
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":73194
      Top             =   4470
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   42
      Left            =   4275
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":74E5E
      Top             =   4470
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   43
      Left            =   3540
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":76B28
      Top             =   4470
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   44
      Left            =   2805
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":787F2
      Top             =   4470
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   45
      Left            =   2070
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":7A4BC
      Top             =   4470
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   46
      Left            =   1335
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":7C186
      Top             =   4470
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   47
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":7DE50
      Top             =   4470
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   32
      Left            =   5760
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":7FB1A
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   33
      Left            =   5010
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":817E4
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   34
      Left            =   4275
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":834AE
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   35
      Left            =   3540
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":85178
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   36
      Left            =   2805
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":86E42
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   37
      Left            =   2070
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":88B0C
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   38
      Left            =   1335
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":8A7D6
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image Board 
      Height          =   720
      Index           =   39
      Left            =   600
      MousePointer    =   99  'Custom
      Picture         =   "RJSoftChess.frx":8C4A0
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label CommandMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONNECT"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Index           =   3
      Left            =   8610
      TabIndex        =   20
      Top             =   915
      Width           =   840
   End
   Begin VB.Image CommandImage 
      Height          =   420
      Index           =   3
      Left            =   8370
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label CommandMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "END"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   7260
      TabIndex        =   17
      Top             =   915
      Width           =   840
   End
   Begin VB.Label CommandMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RESIGN"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Index           =   1
      Left            =   8610
      TabIndex        =   18
      Top             =   465
      Width           =   840
   End
   Begin VB.Label CommandMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Index           =   0
      Left            =   7260
      TabIndex        =   19
      Top             =   465
      Width           =   840
   End
   Begin VB.Image CommandImage 
      Height          =   420
      Index           =   0
      Left            =   7020
      Picture         =   "RJSoftChess.frx":8E16A
      Top             =   420
      Width           =   1305
   End
   Begin VB.Image CommandImage 
      Height          =   420
      Index           =   1
      Left            =   8370
      Top             =   420
      Width           =   1305
   End
   Begin VB.Image CommandImage 
      Height          =   420
      Index           =   2
      Left            =   7020
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label LabelY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   6
      Left            =   45
      TabIndex        =   7
      Top             =   5385
      Width           =   225
   End
   Begin VB.Label LabelY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   5
      Left            =   45
      TabIndex        =   6
      Top             =   4650
      Width           =   225
   End
   Begin VB.Label LabelY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   4
      Left            =   45
      TabIndex        =   5
      Top             =   3900
      Width           =   225
   End
   Begin VB.Label LabelY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   3
      Left            =   45
      TabIndex        =   4
      Top             =   3165
      Width           =   225
   End
   Begin VB.Label LabelY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   2
      Left            =   45
      TabIndex        =   3
      Top             =   2430
      Width           =   225
   End
   Begin VB.Label LabelY 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   1695
      Width           =   225
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   5925
      TabIndex        =   16
      Top             =   6975
      Width           =   345
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   5190
      TabIndex        =   15
      Top             =   6975
      Width           =   345
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   2
      Left            =   4455
      TabIndex        =   14
      Top             =   6975
      Width           =   345
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   13
      Top             =   6975
      Width           =   345
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   4
      Left            =   2985
      TabIndex        =   12
      Top             =   6975
      Width           =   345
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   5
      Left            =   2250
      TabIndex        =   11
      Top             =   6975
      Width           =   345
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   6
      Left            =   1515
      TabIndex        =   10
      Top             =   6975
      Width           =   345
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   7
      Left            =   780
      TabIndex        =   9
      Top             =   6975
      Width           =   345
   End
   Begin VB.Image ChessBoard 
      Height          =   6405
      Left            =   360
      MouseIcon       =   "RJSoftChess.frx":8EA44
      MousePointer    =   99  'Custom
      Top             =   510
      Width           =   6360
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "READY FOR NEW GAME."
      BeginProperty Font 
         Name            =   "Vixar ASCI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1380
      TabIndex        =   45
      Top             =   30
      Width           =   5445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   6720
      X2              =   7920
      Y1              =   3705
      Y2              =   3690
   End
End
Attribute VB_Name = "RJSoftChess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Board_Setup(XA As Boolean)

Dim X As Integer, Y As Integer, Z As Integer, ZBoardTempA_H As String, ZBoardTemp1_8 As String

'Set Board and Labels
If ZRotated = True Then
    For X = 0 To 7
        LabelY(X).Caption = CStr(X + 1)
    Next X
    For X = 0 To 7
        LabelX(X).Caption = Chr(65 + X)
    Next X
    Y = 0: Z = 0
    For X = 0 To 63
        Board(X).MouseIcon = MouseImage(0).Picture
        ZBoardTempA_H = Chr(Y + 65)
        ZBoardTemp1_8 = CStr(Z + 1)
        Board(X).DataField = ZBoardTempA_H & ZBoardTemp1_8
        Y = Y + 1
        If Y >= 8 Then
            Y = 0
            Z = Z + 1
            If Z >= 8 Then Z = 0
        End If
    Next X
 Else
    For X = 0 To 7
        LabelY(X).Caption = CStr(7 - X + 1)
    Next X
    For X = 0 To 7
        LabelX(X).Caption = Chr(72 - X)
    Next X
    Y = 0: Z = 0
    For X = 0 To 63
        Board(X).MouseIcon = MouseImage(0).Picture
        ZBoardTempA_H = Chr(72 - Y)
        ZBoardTemp1_8 = CStr(7 - Z + 1)
        Board(X).DataField = ZBoardTempA_H & ZBoardTemp1_8
        Y = Y + 1
        If Y >= 8 Then
            Y = 0
            Z = Z + 1
            If Z >= 8 Then Z = 0
        End If
    Next X
End If

For X = 16 To 47
    Board(X).Picture = Master_Blank.Picture
    Board(X).Tag = ""
Next X

If XA Then
    'White on top
    Board(7).Picture = Master_White(2).Picture
    Board(7).Tag = "WRK"
    Board(0).Picture = Master_White(2).Picture
    Board(0).Tag = "WRK"
    
    Board(6).Picture = Master_White(3).Picture
    Board(6).Tag = "WKN"
    Board(1).Picture = Master_White(3).Picture
    Board(1).Tag = "WKN"
    
    Board(5).Picture = Master_White(4).Picture
    Board(5).Tag = "WBP"
    Board(2).Picture = Master_White(4).Picture
    Board(2).Tag = "WBP"
    
    Board(3).Picture = Master_White(1).Picture
    Board(3).Tag = "WQU"
    
    Board(4).Picture = Master_White(0).Picture
    Board(4).Tag = "WKK"
    
    For X = 8 To 15
        Board(X).Picture = Master_White(5).Picture
        Board(X).Tag = "WPN"
    Next X
    
    'Black on bottom
    Board(63).Picture = Master_Black(2).Picture
    Board(63).Tag = "BRK"
    Board(56).Picture = Master_Black(2).Picture
    Board(56).Tag = "BRK"
    
    Board(62).Picture = Master_Black(3).Picture
    Board(62).Tag = "BKN"
    Board(57).Picture = Master_Black(3).Picture
    Board(57).Tag = "BKN"
    
    Board(61).Picture = Master_Black(4).Picture
    Board(61).Tag = "BBP"
    Board(58).Picture = Master_Black(4).Picture
    Board(58).Tag = "BBP"
    
    Board(59).Picture = Master_Black(1).Picture
    Board(59).Tag = "BQU"
    
    Board(60).Picture = Master_Black(0).Picture
    Board(60).Tag = "BKK"
    
    For X = 48 To 55
        Board(X).Picture = Master_Black(5).Picture
        Board(X).Tag = "BPN"
    Next X
 Else
    'White on bottom
    Board(7).Picture = Master_Black(2).Picture
    Board(7).Tag = "BRK"
    Board(0).Picture = Master_Black(2).Picture
    Board(0).Tag = "BRK"
    
    Board(6).Picture = Master_Black(3).Picture
    Board(6).Tag = "BKN"
    Board(1).Picture = Master_Black(3).Picture
    Board(1).Tag = "BKN"
    
    Board(5).Picture = Master_Black(4).Picture
    Board(5).Tag = "BBP"
    Board(2).Picture = Master_Black(4).Picture
    Board(2).Tag = "BBP"
    
    Board(4).Picture = Master_Black(1).Picture
    Board(4).Tag = "BQU"
    
    Board(3).Picture = Master_Black(0).Picture
    Board(3).Tag = "BKK"
    
    For X = 8 To 15
        Board(X).Picture = Master_Black(5).Picture
        Board(X).Tag = "BPN"
    Next X
    
    'Black on top
    Board(63).Picture = Master_White(2).Picture
    Board(63).Tag = "WRK"
    Board(56).Picture = Master_White(2).Picture
    Board(56).Tag = "WRK"
    
    Board(62).Picture = Master_White(3).Picture
    Board(62).Tag = "WKN"
    Board(57).Picture = Master_White(3).Picture
    Board(57).Tag = "WKN"
    
    Board(61).Picture = Master_White(4).Picture
    Board(61).Tag = "WBP"
    Board(58).Picture = Master_White(4).Picture
    Board(58).Tag = "WBP"
    
    Board(60).Picture = Master_White(1).Picture
    Board(60).Tag = "WQU"
    
    Board(59).Picture = Master_White(0).Picture
    Board(59).Tag = "WKK"
    
    For X = 48 To 55
        Board(X).Picture = Master_White(5).Picture
        Board(X).Tag = "WPN"
    Next X
End If
WhiteCanCastleKingSide = True
WhiteCanCastleQueenSide = True
WhiteKingHasMoved = False
BlackCanCastleKingSide = True
BlackCanCastleQueenSide = True
BlackKingHasMoved = False

End Sub

Private Sub Board_Click(Index As Integer)

If OpInProgress = True Or ZMovingPiece = True Then Exit Sub

ZMovePiece Index, True

End Sub

Private Sub Board_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

DoEvents

If FromSquare.Visible = False And ToSquare.Visible = False Then
    If PreSquare.Top <> Board(Index).Top Then PreSquare.Top = Board(Index).Top
    If PreSquare.Left <> Board(Index).Left Then PreSquare.Left = Board(Index).Left
    If PreSquare.Visible = False Then PreSquare.Visible = True
End If

If ZMovingPiece = True Then Exit Sub

BoardNumber.Caption = CStr(Index)
BoardLocation.Caption = Board(Index).DataField

If FromSquare.Visible = True Then
    If ToSquare.Top <> Board(Index).Top Then ToSquare.Top = Board(Index).Top
    If ToSquare.Left <> Board(Index).Left Then ToSquare.Left = Board(Index).Left
    If ToSquare.Visible = False Then ToSquare.Visible = True
End If

DoEvents

End Sub


Private Sub ChessBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If PreSquare.Visible = True Then PreSquare.Visible = False

End Sub


Private Sub CommandImage_Click(Index As Integer)

CommandMenu_Click Index

End Sub

Private Sub CommandMenu_Click(Index As Integer)
On Local Error Resume Next
Err.Clear

If OpInProgress2 = True Then Exit Sub

OpInProgress2 = True
DoEvents

Dim X As Integer, A As String
Dim szFilename As String, ZTempAllLoaded As Boolean

If Index = 0 Then ZNewGame True

If Index = 1 And ZGameInProcess = True Then ZResignedGame True

If Index = 2 Then
    'Send system message to disconnect from server
    ZSend = "s" & Indicator & "CoNnEcTcAnCeL" & Indicator & Wsck.LocalHostName
    SendOutIP
    Wsck.Close
    DoEvents
    Unload RJSoftChess
    End
End If

If Index = 3 Then
    tMain.Text = ""
    If CommandMenu(3).Caption = "CONNECT" Then
        If Trim(tHost.Text) = "" Then
            MsgBox ("Please make sure a Host has been entered!")
            'Put blinker in host text box
            tHost.SetFocus
            CommandMenu(3).Caption = "CONNECT"
            ConnectLabel.Caption = "NOT CONNECTED"
            tHost.Locked = False
            LocalIP.Locked = False
            OpInProgress2 = False
            Exit Sub
        Else
            ConnectLabel.Caption = "CONNECTING"
            DoEvents
            Randomize
            'Set the IP or Host Computer to connect to
            Wsck.RemoteHost = tHost.Text
            'Randomize a Port setting
            Wsck.LocalPort = Int((9999 * Rnd) + 1)
            'Set the Port to connect to
            Wsck.RemotePort = 2372
            'Connect!
            Wsck.Bind
            'Send system request to connect
            ZSend = "s" & Indicator & "CoNnEcTrEqUeSt" & Indicator & Wsck.LocalHostName
            Wsck.SendData ZSend
            CommandMenu(3).Caption = "HANGUP"
            tHost.Locked = True
            LocalIP.Locked = True
            tSend.SetFocus
        End If
     Else
        'Send system message to disconnect from server
        ZSend = "s" & Indicator & "CoNnEcTcAnCeL" & Indicator & Wsck.LocalHostName
        SendOutIP
        ConnectLabel.Caption = "NOT CONNECTED"
        CommandMenu(3).Caption = "CONNECT"
        tHost.Locked = False
        LocalIP.Locked = False
        txtMaxGameTime.Locked = False
        CommandMenu(1).Enabled = False
        CommandImage(1).Enabled = False
        CommandMenu(0).Enabled = True
        CommandImage(0).Enabled = True
        MoveList.Visible = False
        MoveList.Clear
        If PlayingOption(0).Value = True Then
            PlayingOption(0).Enabled = True
            PlayingOption(1).Enabled = True
         Else
            PlayingOption(1).Enabled = True
            PlayingOption(0).Enabled = True
        End If
        'Remove clients from your collections
        For X = 0 To RJSoftChess.lName.ListCount - 1
            'Select each IP
            RJSoftChess.lName.ListIndex = X
            'Set IP and Port to send to
            RmIP = RJSoftChess.lName.Text
            Client.Remove (RmIP)
            Names.Remove (RmIP)
        Next
        lName.Clear
        ZGuestMode = False
        GuestBoard.Visible = False
        ZGameInProcess = False
        lblStatus.Caption = "READY FOR NEW GAME."
    End If
End If

If Index = 4 Then
    If ZTempPath = "" Then ZTempPath = App.Path
    szFilename = DialogFile(Me.hWnd, 1, "Open", "Board*.bmp", "BitMaps" & Chr(0) & "Board*.bmp" & Chr(0) & "All files" & Chr(0) & "*.*", ZTempPath, "bmp", ZTempPath)
    If szFilename <> "" Then ChessBoard.Picture = LoadPicture(szFilename)
End If

OpInProgress2 = False

End Sub

Private Sub CursorArrow_Click(Index As Integer)

If ZGameInProcess = True Then Exit Sub

If Index = 0 Then
    If Val(txtMaxGameTime.Text) + 1 > 60 Then
        txtMaxGameTime.Text = "60"
     Else
        txtMaxGameTime.Text = CStr(Val(txtMaxGameTime.Text) + 1)
    End If
End If
If Index = 1 Then
    If Val(txtMaxGameTime.Text) - 1 < 1 Then
        txtMaxGameTime.Text = "1"
     Else
        txtMaxGameTime.Text = CStr(Val(txtMaxGameTime.Text) - 1)
    End If
End If

End Sub

Private Sub Form_Load()
On Local Error Resume Next
Err.Clear

Dim X As Integer

For X = 1 To 4
    CommandImage(X).Picture = CommandImage(0).Picture
Next X

ChessBoard.Picture = LoadPicture(App.Path + "\Board_Stone5.bmp")

Wsck.Protocol = sckUDPProtocol
'Set your constant port (must be the same in clients)
Wsck.LocalPort = 2372
'Start listening
Wsck.Bind

'Add the server to the name list
'This would allow you to make a list box in the client that could
'receive all of the names of the people in the room.
RmIP = Wsck.LocalIP
RmPt = 2372
Names.Add Key:=RmIP, Item:="Server"
'Display your IP Address for client use, and Computer Name for network use.
LocalIPLabel.Caption = "LOCAL: " & RmIP & " / " & Wsck.LocalHostName
LocalIP.Text = RmIP

ZRotated = False
Board_Setup ZRotated

DoEvents
OpInProgress = False

End Sub


Private Sub CommandMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CommandMenu(Index).ForeColor = CommandMenu(Index).ForeColor Xor &HFFFFFF
CommandMenu(Index).LinkItem = CommandMenu(Index).ForeColor
DoEvents

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If PreSquare.Visible = True Then PreSquare.Visible = False

End Sub

Private Sub GameTimer_Timer()

If ZGameInProcess = False Or GameTimer.Interval = 0 Then Exit Sub

Dim ZTempTime As String, ZTempTime2 As String, ZTempTime3 As String

If ZRotated = True Then
    If ZWhosTurn = CWhite Then
        ZTempTime = GameTime(1).Caption
     Else
        ZTempTime = GameTime(2).Caption
    End If
 Else
    If ZWhosTurn = CWhite Then
        ZTempTime = GameTime(2).Caption
     Else
        ZTempTime = GameTime(1).Caption
    End If
End If

If GameTime(1).Caption = "0:00" Or GameTime(2).Caption = "0:00" Then Exit Sub

If Len(ZTempTime) = 4 Then ZTempTime = "0" & ZTempTime
ZTempTime2 = Format(TimeSerial(1, Val(Left(ZTempTime, 2)), Val(Right(ZTempTime, 2) - 1)), "HH:MM:SS")
ZTempTime3 = Right(ZTempTime2, 5)
If Left(ZTempTime3, 1) = "0" Then ZTempTime3 = Right(ZTempTime3, 4)

If ZRotated = True Then
    If ZWhosTurn = CWhite Then
        GameTime(1).Caption = ZTempTime3
        If ZTempTime3 = "0:00" Then
            ZResignedGame True
            lblStatus.Caption = "WHITE LOOSES TO TIME."
        End If
     Else
        GameTime(2).Caption = ZTempTime3
        If ZTempTime3 = "0:00" Then
            ZResignedGame True
            lblStatus.Caption = "BLACK LOOSES TO TIME."
        End If
    End If
 Else
    If ZWhosTurn = CWhite Then
        GameTime(2).Caption = ZTempTime3
        If ZTempTime3 = "0:00" Then
            ZResignedGame True
            lblStatus.Caption = "WHITE LOOSES TO TIME."
        End If
     Else
        GameTime(1).Caption = ZTempTime3
        If ZTempTime3 = "0:00" Then
            ZResignedGame True
            lblStatus.Caption = "BLACK LOOSES TO TIME."
        End If
    End If
End If
DoEvents

End Sub

Private Sub lName_DblClick()

tHost.Text = lName.Text
tSend.SetFocus

End Sub


Private Sub lName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If PreSquare.Visible = True Then PreSquare.Visible = False

End Sub


Private Sub PlayingOption_Click(Index As Integer)

If OpInProgress2 = True Or ZGameInProcess = True Then Exit Sub

If PlayingOption(0).Value = True Then
    ZRotated = False
 Else
    ZRotated = True
End If
Board_Setup ZRotated

End Sub

Private Sub tMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If PreSquare.Visible = True Then PreSquare.Visible = False

End Sub


Private Sub tSend_KeyPress(KeyAscii As Integer)

If ConnectLabel.Caption = "CONNECTED" And KeyAscii = 13 Then
    KeyAscii = 0
    ZSend = "t" & UCase(Wsck.LocalHostName) & ": " & tSend.Text & Chr(13) & Chr(10)
    Wsck.SendData ZSend
    tSend.Text = ""
    Exit Sub
End If

End Sub


Private Sub tSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If PreSquare.Visible = True Then PreSquare.Visible = False

End Sub


Private Sub Wsck_DataArrival(ByVal bytesTotal As Long)
On Local Error Resume Next
Err.Clear

Dim A As String, ZTempHeading As String, ZTempTime As String

A = Chr(254) + Chr(192) + Chr(164) + App.CompanyName + App.Comments + Chr(171) + Chr(188) + Chr(215) + Chr(143) + Chr(94) + Chr(204) + Chr(222) + Chr(248) + Chr(128) + Chr(126) + Chr(149) + Chr(241) + Chr(33) + Chr(42) + Chr(159) + Chr(165) + Chr(172) + Chr(127) + Chr(223) + Chr(227) + Chr(63) + Chr(135)
    
'Winsock received a message

Dim DATA As String
Dim DATA2 As String
Dim DATA3 As String
Dim DATA4 As String
Dim DATA5 As String
Dim Nam As String
Dim MsgText As String
Dim X As Integer, Y As Integer

'Retreive message in string format
Wsck.GetData DATA, vbString

'Get client's IP and Port
RmIP = Wsck.RemoteHostIP
RmPt = Wsck.RemotePort

If (RmIP = Trim(LocalIP.Text)) Then Exit Sub

'Get first letter of message
DATA2 = Left(DATA, 1)
'Get the rest of the message
DATA = Mid(DATA, 2)

'If the message is a system command:
If DATA2 = "s" Then
    'If a client wants to connect to the room:
    If Left(DATA, 20) = Indicator & "CoNnEcTrEqUeSt" & Indicator Then
        If ZGameInProcess = False Then
            'Check for Dup IP
            For X = 0 To lName.ListCount - 1
                If lName.List(X) = RmIP Then
                    lName.RemoveItem (X)
                    Client.Remove (RmIP)
                    Names.Remove (RmIP)
                End If
            Next X
            For X = 0 To ConnectedList.ListCount - 1
                If ConnectedList.List(X) = RmIP Then
                    ConnectedList.RemoveItem (X)
                End If
            Next X
            'Extract the client NickName from the message
            Nam = Mid(DATA, 21)
            'Add client's IP and Port to your collections
            Client.Add Key:=RmIP, Item:=RmPt
            Names.Add Key:=RmIP, Item:=Nam
            'Add client's IP to the listbox
            lName.AddItem RmIP
            lName.ListIndex = lName.ListCount - 1
            ConnectedList.AddItem RmIP
        End If
        'Get User IP
        If ZGameInProcess = True Then
            'Send Request Denied
            ZSend = "s" & Indicator & "ReQuEsTdEnIeD" & Indicator & Wsck.LocalHostName
            Wsck.SendData ZSend
         Else
            'Send Request Granted
            ZSend = "s" & Indicator & "ReQuEsTgRaNtEd" & Indicator & Wsck.LocalHostName
            Wsck.SendData ZSend
            CommandMenu(0).Enabled = True
            lblStatus.Caption = "WAITING TO START NEW GAME."
        End If
        Exit Sub
     ElseIf Left(DATA, 19) = Indicator & "ReQuEsTdEnIeD" & Indicator Then
        DoEvents
        CommandMenu(3).Caption = "CONNECT"
        tHost.Locked = False
        LocalIP.Locked = False
        OpInProgress2 = False
        ConnectLabel.Caption = "DENIED"
        lblStatus.Caption = "GAME ALREADY IN PROGRESS."
     'If a client wants to disconnect from the room:
     ElseIf Left(DATA, 19) = Indicator & "CoNnEcTcAnCeL" & Indicator Then
        'Loop through listbox and find client's IP
        For X = 0 To ConnectedList.ListCount - 1
            ConnectedList.ListIndex = X
            RmEx = ConnectedList.Text
            'When found, remove IP from listbox
            If RmEx = RmIP Then
                ZResignedGame False
                ConnectedList.RemoveItem (X)
                lblStatus.Caption = RmIP & " IS NO LONGER CONNECTED."
            End If
        Next X
        Exit Sub
     'If a request granted was recieved
     ElseIf Left(DATA, 20) = Indicator & "ReQuEsTgRaNtEd" & Indicator Then
        ConnectLabel.Caption = "CONNECTED"
        'Check for Dup IP
        For X = 0 To lName.ListCount - 1
            If lName.List(X) = RmIP Then
                lName.RemoveItem (X)
                Client.Remove (RmIP)
                Names.Remove (RmIP)
            End If
        Next X
        'Extract the client NickName from the message
        Nam = Mid(DATA, 21)
        'Add client's IP and Port to your collections
        Client.Add Key:=RmIP, Item:=RmPt
        Names.Add Key:=RmIP, Item:=Nam
        'Add client's IP to the listbox
        lName.AddItem RmIP
        lName.ListIndex = lName.ListCount - 1
        ZTempHeading = "WAITING FOR OTHER PLAYER TO CONNECT."
        For X = 0 To ConnectedList.ListCount - 1
            If ConnectedList.List(X) = RmIP Then
                ZTempHeading = "WAITING TO START NEW GAME."
            End If
        Next X
        lblStatus.Caption = ZTempHeading
    End If
 ElseIf DATA2 = "t" Then
    'Check for echo of last message
    If DATA = ZLastMessage Then Exit Sub
    'Add the text message to your room
    If Len(tMain.Text) + Len(DATA) >= 15000 Then tMain.Text = Right(tMain.Text, 500)
    tMain.SelStart = Len(tMain)
    tMain.SelText = DATA
    'Scroll to the bottom of the room
    tMain.SelStart = Len(tMain)
    'Send to others in the room
    ZSend = "t" & DATA
    SendOutIP
    ZLastMessage = DATA
 ElseIf DATA2 = "c" Then
     If Trim(Right(DATA, 15)) = Trim(LocalIP.Text) Then
        ZGuestMode = False
        GuestBoard.Visible = False
     Else
        ZGuestMode = True
        GuestBoard.Visible = True
        GuestBoard.ZOrder 0
    End If
    If Mid(DATA, 4, 5) = "WhItE" Then
        PlayingOption(1).Value = True
        ZNewGame False
     ElseIf Mid(DATA, 4, 5) = "BlAcK" Then
        PlayingOption(0).Value = True
        ZNewGame False
    End If
    If Mid(DATA, 4, 7) = "NeWgAmE" Then ZNewGame False
    If Mid(DATA, 4, 8) = "ReSiGnEd" Then ZResignedGame False
    If Mid(DATA, 4, 5) = "TiMeR" Then txtMaxGameTime.Text = CStr(Val(Mid(DATA, 9, 2)))
    If Mid(DATA, 4, 5) = "TiMe1" Then
        GameTimer.Interval = 0
        DoEvents
        ZTempTime = Mid(DATA, 9, 5)
        If Left(ZTempTime, 1) = "0" Then ZTempTime = Right(ZTempTime, 4)
        If ZRotated = True Then
            If ZWhosTurn = CWhite Then
                GameTime(2).Caption = ZTempTime
             Else
                GameTime(1).Caption = ZTempTime
            End If
         Else
            If ZWhosTurn = CWhite Then
                GameTime(1).Caption = ZTempTime
             Else
                GameTime(2).Caption = ZTempTime
            End If
        End If
        'lblMoving.Visible = True
        'lblStatus.Visible = False
        AQZ_H = Hour(Now)
        AQZ_M = Minute(Now)
        AQZ_S = Second(Now)
        Do
            DoEvents
        Loop Until TimeSerial(Hour(Now), Minute(Now), Second(Now)) >= TimeSerial(AQZ_H, AQZ_M, AQZ_S + 1)
        lblMoving.Visible = False
        lblStatus.Visible = True
        GameTimer.Interval = 1000
    End If
 ElseIf DATA2 = "d" Then
    'Check for echo of last move
    If CStr(Val(Mid(DATA, 4, 2)) Xor 63) & "/" & CStr(Val(Mid(DATA, 7, 2)) Xor 63) = ZLastMove Then Exit Sub
    'Move chess pieces
    ZMovingPiece = True
    lblMoving.Visible = True
    lblStatus.Visible = False
    DoEvents
    ZBoardFrom = Val(Mid(DATA, 4, 2)) Xor 63
    BoardNumber.Caption = CStr(ZBoardFrom)
    BoardLocation.Caption = Board(ZBoardFrom).DataField
    ZMovePiece ZBoardFrom, False
    AQZ_H = Hour(Now)
    AQZ_M = Minute(Now)
    AQZ_S = Second(Now)
    Do
        DoEvents
    Loop Until TimeSerial(Hour(Now), Minute(Now), Second(Now)) >= TimeSerial(AQZ_H, AQZ_M, AQZ_S + 1)
    ZBoardTo = Val(Mid(DATA, 7, 2)) Xor 63
    BoardNumber.Caption = CStr(ZBoardTo)
    BoardLocation.Caption = Board(ZBoardTo).DataField
    ZMovePiece ZBoardTo, False
    AQZ_H = Hour(Now)
    AQZ_M = Minute(Now)
    AQZ_S = Second(Now)
    Do
        DoEvents
    Loop Until TimeSerial(Hour(Now), Minute(Now), Second(Now)) >= TimeSerial(AQZ_H, AQZ_M, AQZ_S + 1)
    ZMovingPiece = False
End If

End Sub
Private Sub CommandMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

CommandMenu(Index).ForeColor = CommandMenu(Index).LinkItem Xor &HFFFFFF
DoEvents

End Sub
Public Sub ZNewGame(XA As Boolean)
On Local Error Resume Next
Err.Clear

Dim A As String, X As Integer, ZTempHost As String

A = Chr(254) + Chr(192) + Chr(164) + App.CompanyName + App.Comments + Chr(171) + Chr(188) + Chr(215) + Chr(143) + Chr(94) + Chr(204) + Chr(222) + Chr(248) + Chr(128) + Chr(126) + Chr(149) + Chr(241) + Chr(33) + Chr(42) + Chr(159) + Chr(165) + Chr(172) + Chr(127) + Chr(223) + Chr(227) + Chr(63) + Chr(135)
ZTempHost = String(15, " ")
LSet ZTempHost = Trim(tHost.Text)
OpInProgress2 = True

FromLabel.Caption = "XX"
ToLabel.Caption = "XX"
txtMaxGameTime.Locked = True
GameTime(1).Caption = txtMaxGameTime.Text & ":00"
GameTime(2).Caption = txtMaxGameTime.Text & ":00"

WhiteInCheck = False
BlackInCheck = False
WhiteCheckMate = False
BlackCheckMate = False

If ConnectLabel.Caption = "CONNECTED" Then
    If XA = True Then
        If PlayingOption(0).Value = True Then
            ZSend = "c" & Indicator & "WhItE" & Indicator & ZTempHost
            SendOutIP
         Else
            ZSend = "c" & Indicator & "BlAcK" & Indicator & ZTempHost
            SendOutIP
        End If
        ZSend = "c" & Indicator & "TiMeR" & Format(Val(txtMaxGameTime.Text), "00") & Indicator & ZTempHost
        SendOutIP
        ZSend = "c" & Indicator & "NeWgAmE" & Indicator & ZTempHost
        SendOutIP
    End If
End If
ZToggle = 0
If PlayingOption(0).Value = True Then
    PlayingOption(1).Enabled = False
    PlayingOption(0).Enabled = False
 Else
    PlayingOption(0).Enabled = False
    PlayingOption(1).Enabled = False
End If
PreSquare.Visible = False
FromSquare.Visible = False
ToSquare.Visible = False
Board_Setup ZRotated
lblStatus.Caption = "WAITING FOR WHITE TO MOVE."
ZWhosTurn = CWhite
If ZRotated = True Then
    WhosTurn(1).Picture = GRG(2).Picture
    WhosTurn(2).Picture = GRG(1).Picture
 Else
    WhosTurn(1).Picture = GRG(1).Picture
    WhosTurn(2).Picture = GRG(2).Picture
End If
CommandMenu(0).Enabled = False
CommandImage(0).Enabled = False
MoveList.Clear
MoveList.Visible = True
MoveList.ZOrder 0
If ConnectLabel.Caption = "CONNECTED" Then
    If ZGuestMode = False Then
        If PlayingOption(0).Value = True Then
            CommandMenu(1).Enabled = True
            CommandImage(1).Enabled = True
        End If
        If PlayingOption(1).Value = True Then
            CommandMenu(1).Enabled = False
            CommandImage(1).Enabled = True
        End If
    End If
 Else
    CommandMenu(1).Enabled = True
    CommandImage(1).Enabled = True
End If
DoEvents
ZGameInProcess = True
OpInProgress2 = False
tSend.SetFocus
DoEvents
GameTimer.Interval = 1000

End Sub

Public Sub ZResignedGame(XA As Boolean)
On Local Error Resume Next
Err.Clear

GameTimer.Interval = 0

Dim A As String

A = Chr(254) + Chr(192) + Chr(164) + App.CompanyName + App.Comments + Chr(171) + Chr(188) + Chr(215) + Chr(143) + Chr(94) + Chr(204) + Chr(222) + Chr(248) + Chr(128) + Chr(126) + Chr(149) + Chr(241) + Chr(33) + Chr(42) + Chr(159) + Chr(165) + Chr(172) + Chr(127) + Chr(223) + Chr(227) + Chr(63) + Chr(135)

If ConnectLabel.Caption = "CONNECTED" Then
    If XA = True Then
        ZSend = "c" & Indicator & "ReSiGnEd" & Indicator & Wsck.LocalHostName
        SendOutIP
    End If
End If
If ZWhosTurn = CWhite Then
    lblStatus.Caption = "WHITE RESIGNED."
 Else
    lblStatus.Caption = "BLACK RESIGNED."
End If

If PlayingOption(0).Value = True Then
    PlayingOption(0).Enabled = True
    PlayingOption(1).Enabled = True
 Else
    PlayingOption(1).Enabled = True
    PlayingOption(0).Enabled = True
End If

WhosTurn(1).Picture = GRG(0).Picture
WhosTurn(2).Picture = GRG(0).Picture
CommandMenu(0).Enabled = True
CommandImage(0).Enabled = True
CommandMenu(1).Enabled = False
CommandImage(1).Enabled = False
MoveList.Visible = False
MoveList.Clear
ZToggle = 0
PreSquare.Visible = False
FromSquare.Visible = False
ToSquare.Visible = False
txtMaxGameTime.Locked = False
ZGuestMode = False
GuestBoard.Visible = False
ZGameInProcess = False

End Sub

Public Sub ZMovePiece(XB As Integer, XA As Boolean)
On Local Error Resume Next
Err.Clear

GameTimer.Interval = 0
Dim A As String, ZTempLocation As String, X As Integer, ZTakenPiece As String, ZKillPiece As String, ZTempLegalMove As Boolean, ZTempTime As String, ZTempHost As String

OpInProgress = True
DoEvents

ZTempHost = String(15, " ")
LSet ZTempHost = Trim(tHost.Text)

If ZWhosTurn <> 0 Then
    If ZWhosTurn = CWhite Then
        If WhiteCheckMate = False Then
            If WhiteInCheck = True Then
                lblStatus.Caption = "WHITE IS IN CHECK.  WAITING FOR WHITE TO MOVE."
             Else
                lblStatus.Caption = "WAITING FOR WHITE TO MOVE."
            End If
         Else
            lblStatus.Caption = "GAME OVER: WHITE IS CHECKMATED!"
        End If
     Else
        If BlackCheckMate = False Then
            If BlackInCheck = True Then
                lblStatus.Caption = "BLACK IS IN CHECK.  WAITING FOR BLACK TO MOVE."
              Else
                lblStatus.Caption = "WAITING FOR BLACK TO MOVE."
            End If
         Else
            lblStatus.Caption = "GAME OVER: BLACK IS CHECKMATED!"
        End If
    End If
End If

If ConnectLabel.Caption = "CONNECTED" And XA = True Then
    If PlayingOption(0).Value = True And ZWhosTurn = CBlack Then
        OpInProgress = False
        Exit Sub
    End If
    If PlayingOption(1).Value = True And ZWhosTurn = CWhite Then
        OpInProgress = False
        Exit Sub
    End If
End If

If ZToggle = 0 Then
    If ZWhosTurn = CWhite Then
        If Left(Board(XB).Tag, 1) <> "W" And Board(XB).Tag <> "" Then
            If ConnectLabel.Caption = "NOT CONNECTED" Then
                lblStatus.Caption = "NOT YOUR TURN."
            End If
            OpInProgress = False
            Exit Sub
        End If
    ElseIf ZWhosTurn = CBlack And Board(XB).Tag <> "" Then
        If Left(Board(XB).Tag, 1) <> "B" Then
            If ConnectLabel.Caption = "NOT CONNECTED" Then
                lblStatus.Caption = "NOT YOUR TURN."
            End If
            OpInProgress = False
            Exit Sub
        End If
    End If
    If ZWhosTurn <> 0 And Board(XB).Tag <> "" Then
        FromSquare.Top = Board(XB).Top
        FromSquare.Left = Board(XB).Left
        PreSquare.Visible = False
        FromSquare.Visible = True
        ZToggle = 1
        BoardFrom = Val(BoardNumber.Caption)
        FromLabel.Caption = Board(XB).DataField
        ToLabel.Caption = ""
        If ZWhosTurn = CWhite Then
            MoveList.AddItem "WHITE: " & FromLabel.Caption & " - "
         Else
            MoveList.AddItem "BLACK: " & FromLabel.Caption & " - "
        End If
        MoveList.ListIndex = MoveList.ListCount - 1
    End If
 Else
    If ZToggle = 1 Then
        BoardTo = Val(BoardNumber.Caption)
        ToSquare.Visible = False
        If BoardFrom = BoardTo Then
            FromSquare.Visible = False
            ZToggle = 0
            BoardFrom = -1
            BoardTo = -1
            FromLabel.Caption = ""
            ToLabel.Caption = ""
         Else
            'Check for Legal Move
            WhatPiece = Right(Board(BoardFrom).Tag, 2)
            OppsPiece.Picture = Board(BoardTo).Picture
            CheckForLegalMove True
            'Check for Check
            CheckForCheck True
            If WhiteInCheck = True Or BlackInCheck = True Then
                Board(BoardTo).Picture = OppsPiece.Picture
                If ZWhosTurn = CWhite Then
                    If TriedToCastle = True Then
                        lblStatus.Caption = "WHITE CAN NOT CASTLE WHILE IN CHECK."
                     Else
                        lblStatus.Caption = "WHITE CAN NOT PLACE KING IN CHECK."
                    End If
                 Else
                    If TriedToCastle = True Then
                        lblStatus.Caption = "BLACK CAN NOT CASTLE WHILE IN CHECK."
                     Else
                        lblStatus.Caption = "BLACK CAN NOT PLACE KING IN CHECK."
                    End If
                End If
                LegalMove = False
            End If
            If LegalMove = True Then
                ZTakenPiece = Trim(Right(Board(BoardTo).Tag, 2))
                ZKillPiece = Trim(Right(Board(BoardFrom).Tag, 2))
                Board(BoardTo).Picture = Board(BoardFrom).Picture
                Board(BoardFrom).Picture = Master_Blank.Picture
                FromSquare.Visible = False
                Board(BoardTo).Tag = Board(BoardFrom).Tag
                If ConnectLabel.Caption = "CONNECTED" Then
                    ZTempLocation = Format(BoardFrom, "00") & "/" & Format(BoardTo, "00")
                    ZLastMove = CStr(BoardFrom) & "/" & CStr(BoardTo)
                    'Send text message
                    ZSend = "d" & Indicator & ZTempLocation & Indicator & Wsck.LocalHostName
                    SendOutIP
                End If
                Board(BoardFrom).Tag = ""
                ZToggle = 0
                BoardFrom = -1
                BoardTo = -1
                ToLabel.Caption = Board(XB).DataField
                MoveList.ListIndex = MoveList.ListCount - 1
                If ZTakenPiece <> "" Then
                    MoveList.List(MoveList.ListIndex) = MoveList.List(MoveList.ListIndex) & ToLabel.Caption & " " & ZKillPiece & " X " & ZTakenPiece
                 Else
                    MoveList.List(MoveList.ListIndex) = MoveList.List(MoveList.ListIndex) & ToLabel.Caption
                End If
                'Check for Check
                CheckForCheck False
                'If WhiteInCheck = True Or BlackInCheck = True Then CheckForMate
                If ZWhosTurn = CWhite Then
                    ZWhosTurn = CBlack
                    If ZRotated = True Then
                        WhosTurn(1).Picture = GRG(1).Picture
                        WhosTurn(2).Picture = GRG(2).Picture
                     Else
                        WhosTurn(1).Picture = GRG(2).Picture
                        WhosTurn(2).Picture = GRG(1).Picture
                    End If
                    If BlackCheckMate = False Then
                        If BlackInCheck = True Then
                            lblStatus.Caption = "BLACK IS IN CHECK.  WAITING FOR BLACK TO MOVE."
                            MoveList.List(MoveList.ListIndex) = MoveList.List(MoveList.ListIndex) & " CHECK"
                         Else
                            lblStatus.Caption = "WAITING FOR BLACK TO MOVE."
                        End If
                     Else
                        lblStatus.Caption = "GAME OVER: BLACK IS CHECKMATED!"
                    End If
                 Else
                    ZWhosTurn = CWhite
                    If ZRotated = True Then
                        WhosTurn(1).Picture = GRG(2).Picture
                        WhosTurn(2).Picture = GRG(1).Picture
                     Else
                        WhosTurn(1).Picture = GRG(1).Picture
                        WhosTurn(2).Picture = GRG(2).Picture
                    End If
                    If WhiteCheckMate = False Then
                        If WhiteInCheck = True Then
                            lblStatus.Caption = "WHITE IS IN CHECK.  WAITING FOR WHITE TO MOVE."
                            MoveList.List(MoveList.ListIndex) = MoveList.List(MoveList.ListIndex) & " CHECK"
                         Else
                            lblStatus.Caption = "WAITING FOR WHITE TO MOVE."
                        End If
                     Else
                        lblStatus.Caption = "GAME OVER: WHITE IS CHECKMATED!"
                    End If
                End If
                DoEvents
                'Send Timer1 message
                If XA = True Then
                    If ZRotated = True Then
                        If ZWhosTurn = CWhite Then
                            ZTempTime = GameTime(2).Caption
                         Else
                            ZTempTime = GameTime(1).Caption
                        End If
                     Else
                        If ZWhosTurn = CWhite Then
                            ZTempTime = GameTime(1).Caption
                         Else
                            ZTempTime = GameTime(2).Caption
                        End If
                    End If
                    If Len(ZTempTime) = 4 Then ZTempTime = "0" & ZTempTime
                    If ConnectLabel.Caption = "CONNECTED" Then
                        AQZ_H = Hour(Now)
                        AQZ_M = Minute(Now)
                        AQZ_S = Second(Now)
                        Do
                            DoEvents
                        Loop Until TimeSerial(Hour(Now), Minute(Now), Second(Now)) >= TimeSerial(AQZ_H, AQZ_M, AQZ_S + 2)
                        ZSend = "c" & Indicator & "TiMe1" & ZTempTime & Indicator & ZTempHost
                        SendOutIP
                    End If
                End If
             Else
                BoardTo = -1
            End If
        End If
    End If
End If
If ConnectLabel.Caption = "CONNECTED" And ZGuestMode = False Then
    If PlayingOption(0).Value = True And ZWhosTurn = CWhite Then
        CommandMenu(1).Enabled = True
        CommandImage(1).Enabled = True
     ElseIf PlayingOption(1).Value = True And ZWhosTurn = CBlack Then
        CommandMenu(1).Enabled = True
        CommandImage(1).Enabled = True
     Else
        CommandMenu(1).Enabled = False
        CommandImage(1).Enabled = False
    End If
End If

If Right(lblStatus.Caption, 11) = "CHECKMATED!" Then
    CommandMenu(0).Enabled = True
    CommandImage(0).Enabled = True
    CommandMenu(1).Enabled = False
    CommandImage(1).Enabled = False
    ZWhosTurn = 0
End If

GameTimer.Interval = 1000
DoEvents
OpInProgress = False

End Sub
