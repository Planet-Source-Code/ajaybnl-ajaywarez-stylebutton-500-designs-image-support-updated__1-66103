VERSION 5.00
Begin VB.Form frmCreate 
   BackColor       =   &H00878F9C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StyleButton Themes Manager 1.1"
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmcreate1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic_tmp1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame frm_prop 
      BackColor       =   &H00878F9C&
      Caption         =   "Style Properties :"
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   37
         Text            =   "0"
         Top             =   2715
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   36
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Text            =   "0"
         Top             =   2715
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   34
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00878F9C&
         BorderStyle     =   0  'None
         Height          =   840
         Left            =   2400
         ScaleHeight     =   840
         ScaleWidth      =   2760
         TabIndex        =   33
         Top             =   4020
         Width           =   2760
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   30
         Left            =   420
         Min             =   15
         SmallChange     =   15
         TabIndex        =   32
         Top             =   5160
         Value           =   15
         Width           =   1335
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1215
         LargeChange     =   30
         Left            =   1740
         Min             =   15
         SmallChange     =   15
         TabIndex        =   31
         Top             =   3960
         Value           =   15
         Width           =   255
      End
      Begin VB.PictureBox Picture9 
         Height          =   1170
         Left            =   360
         ScaleHeight     =   1110
         ScaleWidth      =   1275
         TabIndex        =   29
         Top             =   3960
         Width           =   1335
         Begin VB.PictureBox Picture4 
            AutoSize        =   -1  'True
            BackColor       =   &H00878F9C&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   0
            ScaleHeight     =   975
            ScaleWidth      =   1095
            TabIndex        =   30
            Top             =   0
            Width           =   1095
            Begin VB.Line Line3 
               DrawMode        =   6  'Mask Pen Not
               X1              =   120
               X2              =   120
               Y1              =   0
               Y2              =   720
            End
            Begin VB.Line Line4 
               DrawMode        =   6  'Mask Pen Not
               X1              =   360
               X2              =   360
               Y1              =   0
               Y2              =   720
            End
            Begin VB.Line Line5 
               DrawMode        =   6  'Mask Pen Not
               X1              =   0
               X2              =   480
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line6 
               DrawMode        =   6  'Mask Pen Not
               X1              =   0
               X2              =   480
               Y1              =   480
               Y2              =   480
            End
         End
      End
      Begin VB.PictureBox Picture10 
         Height          =   1455
         Left            =   360
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   27
         Top             =   840
         Width           =   1455
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   0
            Picture         =   "frmcreate1.frx":08D2
            ScaleHeight     =   720
            ScaleWidth      =   1020
            TabIndex        =   28
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.PictureBox Picture11 
         Height          =   1455
         Left            =   2040
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   25
         Top             =   840
         Width           =   1455
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   0
            Picture         =   "frmcreate1.frx":2F54
            ScaleHeight     =   720
            ScaleWidth      =   1020
            TabIndex        =   26
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.PictureBox Picture12 
         Height          =   1455
         Left            =   3720
         ScaleHeight     =   1395
         ScaleWidth      =   1395
         TabIndex        =   23
         Top             =   840
         Width           =   1455
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            Height          =   720
            Left            =   0
            Picture         =   "frmcreate1.frx":55D6
            ScaleHeight     =   720
            ScaleWidth      =   1020
            TabIndex        =   24
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   22
         Text            =   "0"
         Top             =   2715
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   21
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton StyleButton4 
         Caption         =   "Save 3 State Pictures"
         Height          =   375
         Left            =   2400
         TabIndex        =   20
         Top             =   5040
         Width           =   2775
      End
      Begin VB.CommandButton StyleButton1 
         Caption         =   "Save"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton ctlButton3 
         Caption         =   "Back"
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   6000
         Width           =   1335
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         Picture         =   "frmcreate1.frx":7C58
         ScaleHeight     =   615
         ScaleWidth      =   735
         TabIndex        =   13
         Top             =   5640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Timer tmr_label 
         Interval        =   2000
         Left            =   960
         Top             =   5760
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Original Style Properties (Dblclick pictures to change,click to test)"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   360
         TabIndex        =   48
         Top             =   240
         Width           =   4650
      End
      Begin VB.Label Label8 
         BackColor       =   &H00878F9C&
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Index           =   5
         Left            =   3990
         TabIndex        =   47
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00878F9C&
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Index           =   4
         Left            =   3990
         TabIndex        =   46
         Top             =   2415
         Width           =   135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00878F9C&
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Index           =   2
         Left            =   2310
         TabIndex        =   45
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00878F9C&
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Index           =   1
         Left            =   2310
         TabIndex        =   44
         Top             =   2415
         Width           =   135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00878F9C&
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Index           =   3
         Left            =   525
         TabIndex        =   43
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00878F9C&
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Index           =   0
         Left            =   525
         TabIndex        =   42
         Top             =   2415
         Width           =   135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Button State Preview :"
         Height          =   195
         Left            =   3000
         TabIndex        =   41
         Top             =   3720
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "notFocused                    onFocus                        onClick"
         Height          =   195
         Left            =   420
         TabIndex        =   40
         Top             =   600
         Width           =   3915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Cordinates :"
         Height          =   195
         Left            =   360
         TabIndex        =   39
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Shape Shape3 
         Height          =   975
         Left            =   2340
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Fill New Cordinates Manually. This Section is Just for Testing"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   360
         TabIndex        =   38
         Top             =   3360
         Width           =   4830
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   2895
         Left            =   240
         Top             =   240
         Width           =   5175
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   2190
         Left            =   240
         Top             =   3360
         Width           =   5175
      End
   End
   Begin VB.Frame frm_styles 
      BackColor       =   &H00878F9C&
      Caption         =   "Manage Styles :"
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.CheckBox Check2 
         BackColor       =   &H00878F9C&
         Caption         =   "Use Resample"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton ctlButton2 
         Caption         =   "Add Button"
         Height          =   375
         Left            =   4080
         TabIndex        =   19
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton StyleButton3 
         Caption         =   "Open Theme"
         Height          =   375
         Left            =   3960
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton StyleButton2 
         Caption         =   "Create New Theme"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00878F9C&
         Caption         =   "Show Actual Images"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5295
      End
      Begin VB.PictureBox outer_frame 
         BackColor       =   &H00878F9C&
         BorderStyle     =   0  'None
         Height          =   4395
         Left            =   120
         ScaleHeight     =   4395
         ScaleWidth      =   5295
         TabIndex        =   3
         Top             =   1560
         Width           =   5295
         Begin VB.VScrollBar inner_scroll 
            Height          =   4335
            LargeChange     =   1000
            Left            =   5025
            Max             =   0
            SmallChange     =   100
            TabIndex        =   4
            Top             =   45
            Width           =   255
         End
         Begin VB.PictureBox inner_frame 
            BackColor       =   &H00878F9C&
            BorderStyle     =   0  'None
            Height          =   2000
            Left            =   0
            ScaleHeight     =   1995
            ScaleWidth      =   5055
            TabIndex        =   5
            Top             =   120
            Width           =   5055
            Begin VB.PictureBox Shape1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   0
               ScaleHeight     =   345
               ScaleWidth      =   345
               TabIndex        =   12
               Top             =   0
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.PictureBox State1 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00878F9C&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   120
               ScaleHeight     =   300
               ScaleWidth      =   1335
               TabIndex        =   8
               Top             =   120
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.PictureBox State2 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00878F9C&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   1560
               ScaleHeight     =   300
               ScaleWidth      =   1335
               TabIndex        =   7
               Top             =   120
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.PictureBox State3 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00878F9C&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   3000
               ScaleHeight     =   300
               ScaleWidth      =   1335
               TabIndex        =   6
               Top             =   120
               Visible         =   0   'False
               Width           =   1335
            End
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Theme File :"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Styles in File :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1590
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu EditStyle 
         Caption         =   "&EDIT"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu delstyle 
         Caption         =   "&DELETE"
      End
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private CurrentIndex As Long
Private Times                    As Long
Private Skinimages()             As StdPicture
Private destWidth                 As Long    'Button Width':( Missing Scope
Private destHeight                As Long    'Button Height':( Missing Scope
Private xSection               As Long    ':( Missing Scope
Private ySection               As Long    ':( Missing Scope



                                               

Sub DeleteStyle(File As String, StyleNO As Integer)
CheckBusy
blnBusy = True

Dim a1 As Long, A2 As Long, HasStyles As Boolean
'LoadThemeFile File
DoEvents
If UBound(Skinimages) > 0 Then
For a1 = 1 To UBound(Skinimages) Step 3
A2 = A2 + 1
If StyleNO = A2 Then
DoEvents
Else
AddStyletoFile File & "new.xbn", SavePicFileBytes(Skinimages(a1)), SavePicFileBytes(Skinimages(a1 + 1)), SavePicFileBytes(Skinimages(a1 + 2)), lngX(a1) * 15, lngY(a1) * 15, lngX(a1 + 1) * 15, lngY(a1 + 1) * 15, lngX(a1 + 2) * 15, lngY(a1 + 2) * 15
HasStyles = True
End If
Next
End If
Kill File
If HasStyles = True Then
Name File & "new.xbn" As File
End If
blnBusy = False

End Sub


Sub LoadStyleValues(Picture1 As PictureBox, x As Long, Y As Long)
    On Error Resume Next
    If ValidateThemeFile = False Then
        Exit Sub
    End If
    
    Picture4.Picture = Picture1.Picture
    HScroll1.Max = ((Picture1.Width) / 2)
    If x > HScroll1.Max Then
    HScroll1.Value = HScroll1.Max
    Else
    HScroll1.Value = x
    End If
    VScroll1.Max = ((Picture1.Height) / 2)
    If Y > VScroll1.Max Then
    VScroll1.Value = VScroll1.Max
    Else
    VScroll1.Value = Y
    End If
    VScroll1_Change
     err.Clear

    
End Sub

Private Sub AddStyletoFile(ByVal File As String, _
                           ByVal Style1 As String, _
                           ByVal Style2 As String, _
                           ByVal Style3 As String, _
                           ByVal xS1 As String, _
                           ByVal ys1 As String, _
                           ByVal xs2 As String, _
                           ByVal ys2 As String, _
                           ByVal xS3 As String, _
                           ByVal yS3 As String)
CheckBusy
blnBusy = True
  ' This Function Describes How To Extract Data From .xbn File
  
  Dim nData     As String
  Dim nFileData As String

    'Open Theme File
    If LenB(Dir(File)) = 0 Then
        nData = String$(12, 0)
     Else
        nData = String$(FileLen(File), 0)
        Open File For Binary As #1
        Get #1, , nData
        Close #1
    End If
    nFileData = nData & Chr$(255) & Chr$(0) & Chr$(255) & Chr$(0) & Chr$(255) & Chr$(0) & Chr$(255) & Chr$(0) & Chr$(Int(Int(ys1) / 15)) & Chr$(0) & Chr$(0) _
& Chr$(0) & Chr$(Int(Int(xS1) / 15)) & Chr$(0) & Chr$(0) & Chr$(0) & Style1 & Chr$(255) & Chr$(0) & Chr$(255) & Chr$(0) & Chr$(255) & Chr$(0) & Chr$(255) _
& Chr$(0) & Chr$(Int(Int(ys2) / 15)) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(Int(Int(xs2) / 15)) & Chr$(0) & Chr$(0) & Chr$(0) & Style2 & Chr$(255) & Chr$(0) _
& Chr$(255) & Chr$(0) & Chr$(255) & Chr$(0) & Chr$(255) & Chr$(0) & Chr$(Int(Int(yS3) / 15)) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(Int(Int(xS3) / 15)) _
& Chr$(0) & Chr$(0) & Chr$(0) & Style3

    'If Dir(File) <> "" Then Kill File':( --> replaced by:
    If Dir(File) <> "" Then
        Kill File
    End If
    Open File For Binary As #1
    Put #1, , nFileData
    Close #1
    nFileData = ""
    nData = ""

blnBusy = False
Exit Sub
err:
    Close #1
    MsgBox "Error Writing File!"
blnBusy = False
End Sub


Private Function ValidateThemeFile() As Boolean
    If Text1.Text <> "" Then
        If Left$(Text1.Text, 3) Like "*:\" Or Left$(Text1.Text, 3) Like "*:/" Then
            ValidateThemeFile = True
         Else
            MsgBox "Invalid Theme Path or File Name!"
        End If
     Else
        MsgBox "Please Create or Browse a Theme File First!"
    End If
End Function


Private Sub Focused(txtText2 As TextBox)
    txtText2.SelStart = 0
    txtText2.SelLength = Len(txtText2.Text) + 1
End Sub





Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
Check2.Visible = False
Else
Check2.Visible = True
End If
LoadButtons
End Sub

Private Sub Check2_Click()
LoadButtons
End Sub

Private Sub ctlButton2_Click()
If ValidateThemeFile = True Then
RefreshEditView
ShowForm frm_prop
End If
End Sub

Private Sub ctlButton3_Click()
ShowForm frm_styles
End Sub

Private Sub delstyle_Click()
If MsgBox("Do You Really Want to Delete the Style from the File?", vbYesNo, "Delete Style From File") = vbYes Then
DeleteStyle Text1.Text, CurrentIndex + 1
LoadThemeFile Text1.Text, Skinimages
LoadButtons
End If
End Sub

Private Sub EditStyle_Click()
RefreshEditView
LoadButton CurrentIndex + 0
ShowForm frm_prop
End Sub

Private Sub Form_Load()
'SavePicture Picture1.Picture, "c:\A.bmp"

ShowForm frm_styles

End Sub
Sub ShowForm(nfrm As Object)
nfrm.Visible = True
nfrm.ZOrder 0
Width = nfrm.Width + 100
Height = nfrm.Height + 500
End Sub
Private Sub HScroll1_Change()
HScroll1.Value = Int(HScroll1.Value / 15) * 15
    Line3.Y1 = 0
    Line3.Y2 = Picture4.Height
    Line4.Y1 = 0
    Line4.Y2 = Picture4.Height
    Line3.X1 = HScroll1.Value
    Line3.X2 = HScroll1.Value
    Line4.X1 = (Picture4.Width - 20) - HScroll1.Value
    Line4.X2 = (Picture4.Width - 20) - HScroll1.Value
    Label3.Caption = "Cordinates : X=" & HScroll1.Value & " Y=" & VScroll1.Value
    'Label7.Caption = "Status : Try the Button Preview and Click Add After Selecting 3 State Images"
    TryButton

End Sub

Private Sub HScroll1_Scroll()

    HScroll1_Change

End Sub

Private Sub inner_Scroll_Change()

  'Scroll

    inner_frame.Top = inner_scroll.Value

End Sub

Private Sub inner_scroll_Scroll()

  'Scroll

    inner_Scroll_Change

End Sub
Private Sub DrawPicture(ByVal N As Long, _
                        Obj As PictureBox)
   
    xSection = lngX(N) * 15
    ySection = lngY(N) * 15
    If N = 1 Then
    
    destHeight = 500
    destWidth = 1300
    pic_tmp1.Width = destWidth
    pic_tmp1.Height = destHeight
    End If
    Obj.BackColor = Me.BackColor
    Set Obj.Picture = Nothing
    pic_tmp1.Cls
    RenderButton Skinimages(N), pic_tmp1.hdc, xSection, ySection, destWidth, destHeight, CBool(Check2.Value), Me.BackColor
    
    Obj.Picture = pic_tmp1.Image
    
    End Sub


Private Sub LoadButtons()
CheckBusy
RefreshStyles
blnBusy = True
On Error GoTo err

  Dim lngImageIndex As Long
  Dim lngButtonIndex As Long
  Dim A3 As Long

    'Load Buttons
    inner_frame.Visible = False
    Times = timeGetTime
    'Unload Buttons (Hide The First One Because We Cant Unload It Ha Ha Ha)
    State1(0).Visible = False
    State3(0).Visible = False
    State2(0).Visible = False
    
    If State1.Count > 1 Then
        For A3 = 1 To State1.Count - 1
            Unload State1(A3)
            Unload State2(A3)
            Unload State3(A3)
            Next A3
    End If
        DoEvents

    
    'If Skins Found
    
    If UBound(Skinimages) > 0 Then
        ' Set First Button Skin
        lngButtontoStyleIndex(0) = 1
        lngImageIndex = 1
        State1(lngButtonIndex).Visible = True
        State2(lngButtonIndex).Visible = True
        State3(lngButtonIndex).Visible = True
        State1(lngButtonIndex).ZOrder 0
        State2(lngButtonIndex).ZOrder 0
        State3(lngButtonIndex).ZOrder 0
        
        
        With State1(lngButtonIndex)
        .BackColor = Me.BackColor
        If Check1.Value = 1 Then
               
            .Picture = Skinimages(lngImageIndex)
            State2(lngButtonIndex).Picture = Skinimages(lngImageIndex + 1)
            State3(lngButtonIndex).Picture = Skinimages(lngImageIndex + 2)
            Else
            
          DrawPicture lngImageIndex, State1(lngButtonIndex)
        DrawPicture lngImageIndex + 1, State2(lngButtonIndex)
        DrawPicture lngImageIndex + 2, State3(lngButtonIndex)
            
            End If
        End With
        
        'Set All Rest Button Skins
        lngButtonIndex = 0
        
        For lngImageIndex = 4 To UBound(Skinimages) Step 3 ' All Images Are 3 State Buttons
            lngButtonIndex = lngButtonIndex + 1
            lngButtontoStyleIndex(lngButtonIndex) = lngImageIndex
            
            
            Load State1(lngButtonIndex)
            Load State2(lngButtonIndex)
            Load State3(lngButtonIndex)
                               
            With State1(lngButtonIndex)
            .BackColor = Me.BackColor
            If Check1.Value = 1 Then
            .Picture = Skinimages(lngImageIndex)
            State2(lngButtonIndex).Picture = Skinimages(lngImageIndex + 1)
            State3(lngButtonIndex).Picture = Skinimages(lngImageIndex + 2)
            Else
            DrawPicture lngImageIndex, State1(lngButtonIndex)
        DrawPicture lngImageIndex + 1, State2(lngButtonIndex)
        DrawPicture lngImageIndex + 2, State3(lngButtonIndex)
            
            End If
            
            .Visible = True
            State2(lngButtonIndex).Visible = True
            State3(lngButtonIndex).Visible = True
            
            .Top = State1(lngButtonIndex - 1).Top + State1(lngButtonIndex - 1).Height + 100
                State2(lngButtonIndex).Top = .Top
                State3(lngButtonIndex).Top = .Top
                
            .Left = State1(lngButtonIndex - 1).Left
                State2(lngButtonIndex).Left = State2(lngButtonIndex - 1).Left
                State3(lngButtonIndex).Left = State3(lngButtonIndex - 1).Left
                State1(lngButtonIndex).ZOrder 0
                State2(lngButtonIndex).ZOrder 0
                State3(lngButtonIndex).ZOrder 0
            End With 'State1(lngButtonIndex)
            
            
            If (State1(lngButtonIndex).Top + State1(lngButtonIndex).Height) > outer_frame.Height Then
                inner_frame.Height = (State1(lngButtonIndex).Top + State1(lngButtonIndex).Height) + 150
                inner_scroll.Max = (outer_frame.Height + 50) - inner_frame.Height
             Else
                inner_frame.Height = outer_frame.Height
                inner_scroll.Max = 0
            End If
            
        Next lngImageIndex
        
        
     Else ' No Theme Found
RefreshStyles
    End If
    
 ' Reset Selection
State3_Click -1
 Debug.Print (timeGetTime - Times) / 1000 & " Secs"
    Times = 0
    
err:
inner_frame.Visible = True
    err.Clear
 blnBusy = False
End Sub
Sub RefreshStyles()
        State1(0).Visible = False
        State2(0).Visible = False
        State3(0).Visible = False
Shape1.Visible = False
CurrentIndex = -1

End Sub
Sub RefreshEditView()
Picture1.Picture = Picture6.Picture
Picture2.Picture = Picture6.Picture
Picture3.Picture = Picture6.Picture
Text2.Text = "0"
Text3.Text = "0"
Text4.Text = "0"
Text5.Text = "0"
Text6.Text = "0"
Text7.Text = "0"
VScroll1.Max = 0
VScroll1.Value = 0
HScroll1.Max = 0
HScroll1.Value = 0
Picture5.Picture = Nothing
Picture4.Picture = Nothing
End Sub
Private Sub Picture1_Click()
LoadStyleValues Picture1, Val(Text2), Val(Text3)
End Sub

Private Sub Picture1_DblClick()

  Dim lngImageIndex As String

    On Error Resume Next
    lngImageIndex = DialogFile(0, 1, "Open Image", vbNullString, "All Files" & Chr$(0) & "*.*", vbNullString, vbNullString)
    If lngImageIndex <> "" Then
        Picture1.Picture = LoadPicture(lngImageIndex)
        'Hilite Picture1
        Text2.Text = ((Picture1.Width - 40) / 2) - 20
        Text3.Text = 15
        Picture1_Click
        'If Err.Number > 0 Then Err.Clear: 'Label7.Caption = "Status : Error in Picture or Values" Else 'Label7.Caption = "Status : Picture Loaded! Please Adjust The Scrollbars For Proper Design of Button!"':( --> replaced by:
        If err.Number > 0 Then
            err.Clear
            'Label7.Caption = "Status : Error in Picture or Values"
         Else
            'Label7.Caption = "Status : Picture Loaded! Please Adjust The Scrollbars For Proper Design of Button!"
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub Picture10_Click()
Picture1_Click
End Sub

Private Sub Picture11_Click()
Picture2_Click
End Sub

Private Sub Picture12_Click()
Picture3_Click
End Sub

Private Sub Picture2_Click()
LoadStyleValues Picture2, Val(Text4), Val(Text5)
End Sub

Private Sub Picture2_DblClick()

  Dim lngImageIndex As String

    lngImageIndex = DialogFile(0, 1, "Open Image", vbNullString, "All Files" & Chr$(0) & "*.*", vbNullString, vbNullString)
    If lngImageIndex <> "" Then
        Picture2.Picture = LoadPicture(lngImageIndex)
        'Hilite Picture2
        Text4.Text = ((Picture2.Width - 40) / 2) - 20
        Text5.Text = 15
        Picture2_Click
        'If Err.Number > 0 Then Err.Clear: 'Label7.Caption = "Status : Error in Picture or Values" Else 'Label7.Caption = "Status : Picture Loaded! Please Adjust The Scrollbars For Proper Design of Button!"':( --> replaced by:
        If err.Number > 0 Then
            err.Clear
            'Label7.Caption = "Status : Error in Picture or Values"
         Else
            'Label7.Caption = "Status : Picture Loaded! Please Adjust The Scrollbars For Proper Design of Button!"
        End If
    End If

End Sub

Private Sub Picture3_Click()
LoadStyleValues Picture3, Val(Text6), Val(Text7)
End Sub

Private Sub Picture3_DblClick()

  Dim lngImageIndex As String

    lngImageIndex = DialogFile(0, 1, "Open Image", vbNullString, "All Files" & Chr$(0) & "*.*", vbNullString, vbNullString)
    If lngImageIndex <> "" Then
        Picture3.Picture = LoadPicture(lngImageIndex)
'        Hilite Picture3
        Text6.Text = ((Picture3.Width - 40) / 2) - 20
        Text7.Text = 15
        Picture3_Click
        'If Err.Number > 0 Then Err.Clear: 'Label7.Caption = "Status : Error in Picture or Values" Else 'Label7.Caption = "Status : Picture Loaded! Please Adjust The Scrollbars For Proper Design of Button!"':( --> replaced by:
        If err.Number > 0 Then
            err.Clear
            'Label7.Caption = "Status : Error in Picture or Values"
         Else
            'Label7.Caption = "Status : Picture Loaded! Please Adjust The Scrollbars For Proper Design of Button!"
        End If
    End If

End Sub



Private Function SavePicFileBytes(Pic As StdPicture)
CheckBusy
blnBusy = True

 Dim D As String
    SavePicture Pic, "tmp.bmp"
    DoEvents
    Open "tmp.bmp" For Binary As #1
    D = String$(LOF(1), 0)
    Get #1, , D
    Close #1
    SavePicFileBytes = D
    If Dir("tmp.bmp") <> "" Then
        Kill "tmp.bmp"
    End If
blnBusy = False

Exit Function
err:
    Close #1
    If Dir("tmp.bmp") <> "" Then
        Kill "tmp.bmp"
    End If
    MsgBox "error Read Picture File"
    blnBusy = False
End Function

Private Sub Shape1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
State3_MouseUp CurrentIndex + 0, Button, Shift, x, Y
End Sub

Private Sub State1_Click(Index As Integer)

    State3_Click Index

End Sub

Private Sub State1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'State3_MouseUp Index, Button, Shift, X, Y

End Sub

Private Sub State1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
State3_MouseUp CurrentIndex + 0, Button, Shift, x, Y
End Sub

Private Sub State2_Click(Index As Integer)

    State3_Click Index

End Sub

Private Sub State2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'State3_MouseUp CurrentIndex + 0, Button, Shift, X, Y

End Sub

Private Sub State2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
State3_MouseUp CurrentIndex + 0, Button, Shift, x, Y
End Sub

Private Sub State3_Click(Index As Integer)
On Error Resume Next
DoEvents
  CurrentIndex = Index
If Index < 0 Then
Shape1.Visible = False
Exit Sub
End If
With Shape1
        .Visible = True
        .Left = State1(Index).Left - 30
        .Top = State1(Index).Top - 30
        .Width = State3(Index).Left + State3(Index).Width - 30
        .Height = State1(Index).Height + 60
        .Refresh
        
 End With 'Shape1
  If err.Number > 0 Then err.Clear
  DoEvents
End Sub

Sub LoadButton(Index As Integer)
  Picture1.Picture = Skinimages(lngButtontoStyleIndex(Index))
    Picture2.Picture = Skinimages(lngButtontoStyleIndex(Index) + 1)
    Picture3.Picture = Skinimages(lngButtontoStyleIndex(Index) + 2)
    
    Text2.Text = lngX(lngButtontoStyleIndex(Index)) * 15
    Text3.Text = lngY(lngButtontoStyleIndex(Index)) * 15
    Text4.Text = lngX(lngButtontoStyleIndex(Index) + 1) * 15
    Text5.Text = lngY(lngButtontoStyleIndex(Index) + 1) * 15
    Text6.Text = lngX(lngButtontoStyleIndex(Index) + 2) * 15
    Text7.Text = lngY(lngButtontoStyleIndex(Index) + 2) * 15
    Picture1_Click
End Sub


Private Sub State3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'State3_MouseUp CurrentIndex + 0, Button, Shift, X, Y

End Sub

Private Sub State3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

State3_Click Index
DoEvents
DoEvents


If Button = 2 Then
If CurrentIndex >= 0 And ValidateThemeFile Then
PopupMenu mnu
End If
End If
End Sub

Private Sub StyleButton1_Click()

    If Left$(StyleButton1.Caption, 4) = "Save" Then
        If Int(Text2.Text) > 0 And Int(Text3.Text) > 0 And Int(Text4.Text) > 0 And Int(Text5.Text) > 0 And Int(Text6.Text) > 0 And Int(Text7.Text) > 0 Then
            AddStyletoFile Text1.Text, SavePicFileBytes(Picture1.Picture), SavePicFileBytes(Picture2.Picture), SavePicFileBytes(Picture3.Picture), Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text, Text7.Text
            'Label7.Caption = "Status : " & "Style Added Sucessfully!"
            LoadThemeFile Text1.Text, Skinimages
            LoadButtons
            MsgBox "Button Sucessfully Added To Theme File.", vbInformation
ShowForm frm_styles
         Else
            MsgBox "Please Ensure That No Cordinate Value Could Be Less Than 15 !"
        End If
    End If

End Sub

Private Sub StyleButton2_Click()

  Dim lngImageIndex As String

    lngImageIndex = DialogFile(0, 2, "Create New Theme File", "Theme1", "Button Theme Files" & Chr$(0) & "*.xbn" & Chr$(0) & "All Files" & Chr$(0) & "*.*", GetSetting(App.Title, "paths", "Newtheme", vbNullString), "xbn")
    'If lngImageIndex <> "" Then SaveSetting App.Title, "paths", "Newtheme", lngImageIndex':( --> replaced by:
    If lngImageIndex <> "" Then
        SaveSetting App.Title, "paths", "Newtheme", lngImageIndex
    End If
    If lngImageIndex <> "" Then
        Text1.Text = lngImageIndex
        LoadThemeFile lngImageIndex, Skinimages
        LoadButtons
        'Label7.Caption = "Status : " & "You can add styles in this New Theme File!"
    End If

End Sub

Private Sub StyleButton3_Click()

  Dim lngImageIndex As String
    lngImageIndex = DialogFile(0, 1, "Open Theme File", vbNullString, "Button Theme Files" & Chr$(0) & "*.xbn" & Chr$(0) & "All Files" & Chr$(0) & "*.*", GetSetting(App.Title, "paths", "opentheme", vbNullString), vbNullString)
    If lngImageIndex <> "" Then
        SaveSetting App.Title, "paths", "opentheme", lngImageIndex
    End If
    If lngImageIndex <> "" Then
        Text1.Text = lngImageIndex
        LoadThemeFile lngImageIndex, Skinimages
        LoadButtons
        If State3(0).Visible Then
            State3_Click 0
        End If
        If State1.Count = 1 And State1(0).Visible = False Then
            'Label7.Caption = "Status : " & "This Theme File has no Styles! However you can add styles to it!"
            MsgBox "This Theme File has no Styles! However you can add styles to it!"
         Else
            'Label7.Caption = "Status : " & "You can manage styles in this Theme File!"
        End If
    End If

End Sub

Private Sub StyleButton4_Click()

    SavePicture Picture1.Picture, Environ$("USERPROFILE") & "\desktop\State1.bmp"
    SavePicture Picture2.Picture, Environ$("USERPROFILE") & "\desktop\State2.bmp"
    SavePicture Picture3.Picture, Environ$("USERPROFILE") & "\desktop\State3.bmp"
    MsgBox "Pictures of 3 State Buttons are Saved on Desktop . Modify it and Add it by Double Clicking on Images!"

End Sub

Private Sub Text2_GotFocus()

    Focused Text2

End Sub

Private Sub Text3_GotFocus()

    Focused Text3

End Sub

Private Sub Text4_GotFocus()

    Focused Text4

End Sub

Private Sub Text5_GotFocus()

    Focused Text5

End Sub

Private Sub Text6_GotFocus()
    Focused Text6
End Sub

Private Sub Text7_GotFocus()
    Focused Text7
End Sub

Private Sub TryButton()

    destWidth = Picture5.Width
    destHeight = Picture5.Height
    xSection = HScroll1.Value
    ySection = VScroll1.Value
    Set Picture5.Picture = Nothing
    RenderButton Picture4.Picture, Picture5.hdc, xSection, ySection, destWidth, destHeight, True
End Sub

Private Sub VScroll1_Change()
VScroll1.Value = Int(VScroll1.Value / 15) * 15

    Line5.X1 = 0
    Line5.X2 = Picture4.Width
    Line6.X1 = 0
    Line6.X2 = Picture4.Width
    Line5.Y1 = VScroll1.Value
    Line5.Y2 = VScroll1.Value
    Line6.Y1 = (Picture4.Height - 20) - VScroll1.Value
    Line6.Y2 = (Picture4.Height - 20) - VScroll1.Value
    Label3.Caption = "Cordinates : X=" & HScroll1.Value & " Y=" & VScroll1.Value
    HScroll1_Change
    TryButton
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub


':)Roja's VB Code Fixer V1.1.93 (12/31/2007 2:27:09 PM) 56 + 1040 = 1096 Lines Thanks Ulli for inspiration and lots of code.

