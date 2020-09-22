VERSION 5.00
Begin VB.Form frmBenchmark 
   Caption         =   "Benchmark"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctlButton ctlButton1 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      theme_no        =   3
      Caption         =   "ctlButton1"
      Enabled         =   -1
      ToolbarMode     =   0   'False
      PicState1       =   "frmBenchmark.frx":0000
      PicState2       =   "frmBenchmark.frx":1A7E
      PicState3       =   "frmBenchmark.frx":34FC
      Skin_Cords      =   "1335,375,75,75,75,75,75,75"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label2 
      Caption         =   "Loaded Buttons :"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "StyleButton Benchmark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmBenchmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a1, toq As Long
toq = 10000
On Error GoTo err
Command1.Caption = "Wait..."
For a1 = 1 To toq
Load ctlButton1(a1)
ctlButton1(a1).Visible = True
'ctlButton1(a1).Top = ctlButton1(a1 - 1).Top + 10
'ctlButton1(a1).ZOrder 0
Label3.Caption = a1
Next
Me.Caption = "Benchmarked : " & a1 & " Your Memory Supports Only Buttons :)"
err:
err.Clear
For a1 = ctlButton1.Count - 1 To 1 Step -1
Unload ctlButton1(a1)
Next

Command1.Caption = "Start"
End Sub

