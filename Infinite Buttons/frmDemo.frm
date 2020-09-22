VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "About - StyleButton"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctlButton ctlButton1 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      theme_file_no   =   19
      theme_no        =   6
      Enabled         =   -1
      ToolbarMode     =   0   'False
      PicState1       =   "frmDemo.frx":0000
      PicState2       =   "frmDemo.frx":356E
      PicState3       =   "frmDemo.frx":6ADC
      Skin_Cords      =   "2055,495,75,75,75,75,75,75"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Benchmark"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Theme Manager"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmDemo.frx":A04A
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit




Private Sub Command1_Click()
frmCreate.Show vbModal

End Sub

Private Sub Command2_Click()
frmBenchmark.Show vbModal

End Sub

Private Sub ctlButton2_Click()

End Sub

Private Sub Form_Load()
Show
End Sub
