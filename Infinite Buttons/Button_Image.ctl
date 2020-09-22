VERSION 5.00
Begin VB.UserControl ctlButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   PropertyPages   =   "Button_Image.ctx":0000
   ScaleHeight     =   870
   ScaleWidth      =   2160
   ToolboxBitmap   =   "Button_Image.ctx":0010
   Begin VB.Timer tmr_focus 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   0
   End
   Begin VB.Label demo_caption 
      AutoSize        =   -1  'True
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape focus_rect 
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      Height          =   255
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "ctlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Warning : This code is Free to Use . You Must Give Author Name and Email While Using .
'Author : Ajay Kumar
'Email : ajaybnl@gmail.com



Option Explicit
Private Const LWA_COLORKEY     As Long = &H1
Private Const GWL_EXSTYLE      As Long = (-20)
Private Const WS_EX_LAYERED    As Long = &H80000
Private Const BM_SETSTATE      As Long = &HF3
Private HasPicture             As Boolean
Private State                  As Integer    'State of Buttons ( 0=Unfocus 1=Focus 2=Pressed )

    
Private PicState1              As StdPicture 'Saves Skin Pic
Private PicState2              As StdPicture 'Saves Skin Pic
Private PicState3              As StdPicture 'Saves Skin Pic
Private sCords As String ' Contains DestWidth,DestHeight,sx1,sy1,sx2,sy2,sx3,sy3

Private BtnPicture             As New StdPicture
Private ntheme_file_no         As Integer    ' selection of file
Private ntheme_no              As Integer    ' selection of theme
Private blnToolbarMode         As Boolean
Private HasFocus               As Boolean
Private Skinimages()             As StdPicture
Private SettingsLoaded As Boolean
'Private lngForeColor As Long
Public Event Click()
Public Event KeyDown(Keycode As Integer)



Public Property Get BackColor() As OLE_COLOR
    BackColor = demo_caption.BackColor
End Property


Public Property Let BackColor(oleVal As OLE_COLOR)

    demo_caption.BackColor = oleVal
    PropertyChanged "Backcolor"
    UserControl_Resize
    Draw False, 0

End Property

Public Property Get ButtonPicture() As StdPicture

    If BtnPicture = 0 Or BtnPicture Is Nothing Then
       Set ButtonPicture = BtnPicture
        HasPicture = False
    Else 'NOT BTNPICTURE...
        Set ButtonPicture = BtnPicture
    End If
    

End Property

Public Property Let ButtonPicture(Pic1 As StdPicture)
    Dim srcWidth As Long, srcHeight As Long
    
    srcWidth = Int(((Pic1.Width * 567) / 1000))
    srcHeight = Int(((Pic1.Height * 567) / 1000))
If srcWidth <= 0 Or srcWidth / 15 > 128 Or srcHeight <= 0 Or srcHeight / 15 > 128 Then
MsgBox "Invalid Picture or Invalid Picture Size! Supports only Maximum 128x128 Pictures!", vbCritical, "Error"
Else
    HasPicture = True
    Set BtnPicture = Pic1
    PropertyChanged "BtnPicture"
    Draw False
End If
End Property

Public Property Set ButtonPicture(Pic1 As StdPicture)
    Dim srcWidth As Long, srcHeight As Long
    
    srcWidth = Int(((Pic1.Width * 567) / 1000))
    srcHeight = Int(((Pic1.Height * 567) / 1000))
If srcWidth <= 0 Or srcWidth / 15 > 128 Or srcHeight <= 0 Or srcHeight / 15 > 128 Then
MsgBox "Invalid Picture or Invalid Picture Size! Supports only Maximum 128x128 Pictures!", vbCritical, "Error"
Else
    HasPicture = True
    Set BtnPicture = Pic1
    PropertyChanged "BtnPicture"
    Draw False
End If


End Property

' Get Caption
Public Property Get Caption() As String

    Caption = demo_caption.Caption

End Property

' Set Caption
Public Property Let Caption(ByVal strVal As String)

    demo_caption.Caption = strVal
    DoEvents
    PropertyChanged ("Caption")
       
    Draw False, 0

End Property

Public Sub Draw(Optional Down As Boolean, _
                Optional ByVal lngState As Integer = 0)
                On Error GoTo err
                
CheckBusy , , "Draw"
blnBusy = True

If PicState1 Is Nothing Then GoTo NOPIC
    ' State of Button (0=not fucus,1= focus,2=click)
    If lngState = 0 Then
        Set UserControl.Picture = PicState1
        State = 0
    ElseIf lngState = 1 Then 'NOT LNGSTATE...
        Set UserControl.Picture = PicState2
        State = 1
    ElseIf lngState = 2 Then 'NOT LNGSTATE...
        Set UserControl.Picture = PicState3
        State = 2
    Else 'NOT LNGSTATE...
        State = 0
        Set UserControl.Picture = PicState1
    End If
    
    
NOPIC:
    
    ' No Style
    If PicState1 Is Nothing Then
    
    
        ' State of Button (0=not fucus,1= focus,2=click)
    If lngState = 0 Then

        State = 0
    ElseIf lngState = 1 Then 'NOT LNGSTATE...

        State = 1
    ElseIf lngState = 2 Then 'NOT LNGSTATE...

        State = 2
    Else 'NOT LNGSTATE...
        State = 0
        
    End If

    
    UserControl.Cls
        If Down Then
            With UserControl
                .ForeColor = &H808080
                UserControl.Line (0, 0)-(UserControl.Width, 0)
                UserControl.Line (0, 0)-(0, UserControl.Height)
                .ForeColor = &HE0E0E0
                UserControl.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height)
                UserControl.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15)
            End With 'UserControl
        Else 'DOWN = FALSE/0
            With UserControl
                .ForeColor = &HE0E0E0
                UserControl.Line (0, 0)-(UserControl.Width, 0)
                UserControl.Line (0, 0)-(0, UserControl.Height)
                .ForeColor = &H808080
                UserControl.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height)
                UserControl.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15)
            End With 'UserControl
        End If
    End If
    
    blnBusy = False
    
    If Not blnToolbarMode Then
        RenderImagenLabel Down
    End If

err:
End Sub

Public Property Get Enabled() As Boolean

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal v As Boolean)

    UserControl.Enabled = v
    Draw False
    PropertyChanged "Enabled"

End Property

Public Property Get Font() As StdFont

    Set Font = demo_caption.Font

End Property

Public Property Let Font(ByVal strVal As StdFont)

    Set demo_caption.Font = strVal
    PropertyChanged ("Font")
    Draw False, 0

End Property

' Set Caption
Public Property Set Font(ByVal strVal As StdFont)

    Set demo_caption.Font = strVal
    PropertyChanged ("Font")
    Draw False, 0

End Property

' Get Forecolor
Public Property Get ForeColor() As OLE_COLOR

    ForeColor = demo_caption.ForeColor

End Property

' Set Fore Color
Public Property Let ForeColor(oleVal As OLE_COLOR)

    demo_caption.ForeColor = oleVal
    PropertyChanged "Forecolor"
    Draw
    'RenderImagenLabel

End Property

'Get Mask Color

Public Property Get nHwnd() As Long

    nHwnd = UserControl.hwnd

End Property



' Draw Caption
Private Sub RenderImagenLabel(Optional ByVal Down As Boolean)
Dim Lbl        As String
Dim RightMode  As Boolean
Dim HasCaption As Boolean
Dim srcWidth  As Long
Dim srcHeight As Long
Dim srcHDC As Long
Dim hBmp As Long
Dim hBmpOld As Long
Dim BMP As BITMAP

        
   
    If BtnPicture = 0 Or BtnPicture Is Nothing Then
        HasPicture = False
    Else 'NOT BTNPICTURE...
    'Load Picture
        If BtnPicture.Type = vbPicTypeBitmap Then
        hBmp = BtnPicture.Handle
       srcHDC = CreateCompatibleDC(0&)
        If srcHDC <> 0 Then
         hBmpOld = SelectObject(srcHDC, hBmp)
            If Not GetObject(hBmp, Len(BMP), BMP) <> 0 Then
             HasPicture = False
             End If  'GetObject
       End If
    End If
    
        'Convert Hmetric To Twips
    srcWidth = Int(((BtnPicture.Width * 567) / 1000))
    srcHeight = Int(((BtnPicture.Height * 567) / 1000))

    HasPicture = True
    End If
        

    If Not UserControl.Enabled Then
            UserControl.ForeColor = &H888888
    Else
        UserControl.ForeColor = demo_caption.ForeColor
    End If
    
    With UserControl
        .BackColor = demo_caption.BackColor
        .DrawMode = vbCopyPen
        .DrawStyle = vbSolid
        .Font = demo_caption.Font
        .FontSize = demo_caption.FontSize
        '.ForeColor = demo_caption.ForeColor
        .FontBold = demo_caption.FontBold
    End With 'USERCONTROL
    
    
    
    
    'Usercontrol Settings
    HasCaption = (LenB(demo_caption.Caption) > 0)
    If HasPicture Then
        With UserControl
            If .Height < (srcHeight + 70) Then
                .Height = srcHeight + 70
                .Refresh
            End If
            If .Width < (srcWidth + 70) Then
                .Width = srcWidth + 70
                .Refresh
            End If
            
        End With 'UserControl
        
    End If
    
    'Mode Set
    If HasPicture Then
        If UserControl.Height < (srcHeight + demo_caption.Height + 200) Then
            RightMode = True
        Else 'NOT USERCONTROL.HEIGHT...
            RightMode = False ''
        End If
    End If
    
    ' Start Label Settings
    If HasPicture And HasCaption Then
    
        If RightMode Then
            ' Text Settings (Trimming) Right Mode
            Lbl = demo_caption.Caption
re:
            If demo_caption.Width > ((UserControl.Width - 450) - srcWidth) Then
                If Len(demo_caption.Caption) > 1 Then
                    demo_caption.Caption = Right$(demo_caption.Caption, Len(demo_caption.Caption) - 1)
                    GoTo re
                End If
            End If
            With demo_caption
                If Not Lbl = .Caption Then
                    .Caption = "..." & .Caption
                    .ToolTipText = Lbl
                End If
            End With 'demo_caption
        
        Else ' not Right Mode'RIGHTMODE = FALSE/0
        
            ' Text Settings (Trimming) Top Mode
            Lbl = demo_caption.Caption
re1:
            If demo_caption.Width > UserControl.Width - 250 Then
                demo_caption.Caption = Right$(demo_caption.Caption, Len(demo_caption.Caption) - 1)
                GoTo re1
            End If
            With demo_caption
                If Not Lbl = .Caption Then
                    .Caption = "..." & .Caption
                    .ToolTipText = Lbl
                End If
            End With 'demo_caption
        End If
        
        ' End Label Settings
        demo_caption.ToolTipText = Lbl
        
    ElseIf HasCaption Then 'NOT HASPICTURE...
        ' Text Settings (Trimming) No Picture
        Lbl = demo_caption.Caption
re2:
        If demo_caption.Width > UserControl.Width - 150 And demo_caption.Caption <> "" And Len(demo_caption.Caption) > 3 Then
            demo_caption.Caption = Right$(demo_caption.Caption, Len(demo_caption.Caption) - 1)
            GoTo re2
        End If
        
        With demo_caption
            If Not Lbl = .Caption Then
                .Caption = "..." & .Caption
                .ToolTipText = Lbl
            End If
        End With 'demo_caption
    End If
    
    
    
    ' Draw Picture
    If HasPicture And HasCaption Then
        If RightMode Then
            TransparentBlt UserControl.hDC, 150 / 15, IIf(Down, 2, 0) + ((UserControl.Height - (srcHeight)) / 2) / 15, srcWidth / 15, srcHeight / 15, srcHDC, 0, 0, srcWidth / 15, srcHeight / 15, GetPixel(srcHDC, 0, 0)
        Else 'RIGHTMODE = FALSE/0
            TransparentBlt UserControl.hDC, (IIf(Down, 2, 0) + ((UserControl.Width - srcWidth) / 2)) / 15, IIf(Down, 2, 0) + ((UserControl.Height - (srcHeight + demo_caption.Height + 30)) / 2) / 15, srcWidth / 15, srcHeight / 15, srcHDC, 0, 0, srcWidth / 15, srcHeight / 15, GetPixel(srcHDC, 0, 0)
        End If
    ElseIf HasPicture And Not HasCaption Then  'NOT HASPICTURE...
        TransparentBlt UserControl.hDC, (((UserControl.Width - 20) - (srcWidth)) / 2) / 15, IIf(Down, 2, 0) + (((UserControl.Height) - (srcHeight)) / 2) / 15, srcWidth / 15, srcHeight / 15, srcHDC, 0, 0, srcWidth / 15, srcHeight / 15, GetPixel(srcHDC, 0, 0)
    End If
    
    If HasPicture Then
Call SelectObject(srcHDC, hBmpOld)
DeleteObject hBmpOld
DeleteDC srcHDC

    End If
    
    'Set Lable
    If HasPicture And HasCaption Then
        If RightMode Then
            UserControl.CurrentY = (((UserControl.Height - (demo_caption.Height)) / 2)) + IIf(Down, 15, 0)
            UserControl.CurrentX = srcWidth + (((UserControl.Width - srcWidth) - demo_caption.Width) / 2) + IIf(Down, 10, 0)
        Else 'RIGHTMODE = FALSE/0
            UserControl.CurrentY = (((UserControl.Height - 250) + (((UserControl.Height - (srcHeight + demo_caption.Height + 30)) / 2) + srcHeight)) / 2) + IIf(Down, 15, 0)
            'UserControl.CurrentY = (((UserControl.Height - (srcHeight)) / 2) + srcHeight) + IIf(Down = True, 15, 0)
            UserControl.CurrentX = IIf(Down, 10, 0) + ((UserControl.Width - demo_caption.Width) / 2)
        End If
    Else 'NOT HASPICTURE...
        UserControl.CurrentY = IIf(Down, 15, 0) + ((UserControl.Height - demo_caption.Height) / 2)
        UserControl.CurrentX = IIf(Down, 10, 0) + ((UserControl.Width - demo_caption.Width) / 2)
    End If
    'demo_caption.Left = UserControl.CurrentX
    'demo_caption.Top = UserControl.CurrentY
    
        
    'UserControl.CurrentX = demo_caption.Left
    'UserControl.CurrentY = demo_caption.Top
    With demo_caption
        UserControl.Print .Caption
        .Caption = Lbl
    End With 'demo_caption
    


End Sub

'//////--Set of 3 Functions To Draw The Button--------------------------------------------------------------------------------------------------

' Assign Pictures (Called By Property Page)
' Set Num =0 For First Picture And So On Upto 3 Pics
Public Sub SetSkin(stdState1 As StdPicture, _
                   stdState2 As StdPicture, _
                   stdState3 As StdPicture, _
                   ByVal Cords As String, _
                   ByVal File As Integer, _
                   ByVal Selection As Integer, _
                   Optional InControlCall As Boolean = False)
                   CheckBusy , , "SetSkin"
                   blnBusy = True
    If Not stdState1 Is Nothing Then
        ntheme_file_no = File
        ntheme_no = Selection
        Set PicState1 = stdState1
        Set PicState2 = stdState2
        Set PicState3 = stdState3
        'If You Want to Extract The Skin Pics
If InControlCall = False Then
        PropertyChanged "PicState1"
        PropertyChanged "PicState2"
        PropertyChanged "PicState3"
        PropertyChanged "Skin_Cords"
End If
        sCords = Cords
        
   blnBusy = False
     
        ' Only Fire on New Skin
        If Not InControlCall Then
        ResizenSave True
        End If
        Draw False
        End If
End Sub

Public Property Get theme_file_no() As Integer

    theme_file_no = ntheme_file_no

End Property

Public Property Get theme_no() As Integer

    theme_no = ntheme_no

End Property

' Mouseleave Event Detecter
Private Sub tmr_focus_Timer()

Dim P1           As POINTAPI
Dim PointerHwnd As Long
    GetCursorPos P1
    PointerHwnd = WindowFromPoint(P1.x, P1.Y)
    If Not PointerHwnd = UserControl.hwnd Then
        Draw False, 0
        tmr_focus.Enabled = False
End If
End Sub

Public Property Get ToolbarMode() As Boolean

    ToolbarMode = blnToolbarMode

End Property

Public Property Let ToolbarMode(ByVal i As Boolean)

    blnToolbarMode = i
    Draw False, 0
    PropertyChanged "ToolbarMode"

End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)
Draw
End Sub

Private Sub UserControl_EnterFocus()
HasFocus = True
Draw
End Sub

Private Sub UserControl_ExitFocus()
HasFocus = False
Draw
End Sub

Private Sub UserControl_HitTest(x As Single, Y As Single, HitResult As Integer)
Draw
End Sub

Private Sub UserControl_Initialize()
'First Call
   If SettingsLoaded Then Draw

End Sub

Private Sub UserControl_InitProperties()
demo_caption.Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_KeyDown(Keycode As Integer, _
                                Shift As Integer)

RaiseEvent KeyDown(Keycode)
    
    If Keycode = 13 Then
        RaiseEvent Click
    End If

End Sub

Private Sub UserControl_LostFocus()

    HasFocus = False
    Draw

End Sub

' Draw MouseDown Picture
Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  Y As Single)

CheckBusy , , "Usercontrol_MouseDown"
If Not ToolbarMode Then
        If Button = 1 Then
            Draw True, 2
            tmr_focus.Enabled = False
        End If
    End If

End Sub

' Draw MouseUp Picture if Mouse is Entered
Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  Y As Single)

    If Not ToolbarMode Then
        If State = 0 Then
            Draw False, 1
            tmr_focus.Enabled = True

        End If
        End If

End Sub

' Draw MouseUp Picture if MouseDown is Uccored
Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)

CheckBusy , , "Usercontrol_Mouseup"
If Not ToolbarMode Then
        If State = 2 Then
                tmr_focus.Enabled = True
           If Button = 1 Then
            If x > 0 And x < UserControl.Width And Y > 0 And Y < UserControl.Height Then
Draw False, 1
RaiseEvent Click
            End If
            End If
        End If
    End If
End Sub

' Propertys Handler
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Local Error Resume Next
Dim Val As String
    With PropBag
        Set demo_caption.Font = .ReadProperty("Font", Ambient.Font)
        
        Set PicState1 = .ReadProperty("PicState1", Nothing)
        Set PicState2 = .ReadProperty("PicState2", Nothing)
        Set PicState3 = .ReadProperty("PicState3", Nothing)
        
        Set BtnPicture = .ReadProperty("BtnPicture", Nothing)
        
        ntheme_file_no = .ReadProperty("theme_file_no", 0)
        UserControl.Enabled = CBool(.ReadProperty("Enabled", 0))
        ntheme_no = .ReadProperty("theme_no", 0)
        
        demo_caption.Caption = .ReadProperty("Caption", Ambient.DisplayName)
        demo_caption.ForeColor = .ReadProperty("Forecolor", &H0)
        demo_caption.BackColor = .ReadProperty("Backcolor", &H8000000F)
        blnToolbarMode = .ReadProperty("ToolbarMode", "0")
        
        ' Skin Property's
       sCords = .ReadProperty("Skin_Cords", "0,0,0,0,0,0,0,0")
              
        ' Set Button Skin
        If Not (PicState1 Is Nothing) Then
            SetSkin PicState1, PicState2, PicState3, sCords, ntheme_file_no, ntheme_no, True
        End If
        
        
        
        If BtnPicture Is Nothing Or BtnPicture = 0 Then HasPicture = False Else HasPicture = True
        
    End With 'PROPBAG
    
SettingsLoaded = True

End Sub

' Resize and Reprint Called on Resize Usercontrol , Set Skin , Back color changed
' To Fast The Code , It is Only Called When Changes Applied Related To Skin
Sub ResizenSave(Optional Save As Boolean = False)
CheckBusy , , "Resizensave"
blnBusy = True
DoEvents
    If Not PicState1 Is Nothing Then
        With UserControl
            sCords = .Width & "," & .Height & "," & Val(Split(sCords, ",")(2)) & "," & Val(Split(sCords, ",")(3)) & "," & Val(Split(sCords, ",")(4)) & "," & Val(Split(sCords, ",")(5)) & "," & Val(Split(sCords, ",")(6)) & "," & Val(Split(sCords, ",")(7))
        End With 'UserControl
           
        If Val(Split(sCords, ",")(1)) <= 0 Then
        sCords = 1000 & "," & 300 & "," & Val(Split(sCords, ",")(2)) & "," & Val(Split(sCords, ",")(3)) & "," & Val(Split(sCords, ",")(4)) & "," & Val(Split(sCords, ",")(5)) & "," & Val(Split(sCords, ",")(6)) & "," & Val(Split(sCords, ",")(7))
        End If
        
     
     UserControl.BackColor = demo_caption.BackColor
     
    Set UserControl.Picture = Nothing

        RenderButton PicState1, UserControl.hDC, Val(Split(sCords, ",")(2)), Val(Split(sCords, ",")(3)), Val(Split(sCords, ",")(0)), Val(Split(sCords, ",")(1))
        Set PicState1 = UserControl.Image
        If Save = True Then PropertyChanged "PicState1"
        UserControl.Cls
        
        RenderButton PicState2, UserControl.hDC, Val(Split(sCords, ",")(4)), Val(Split(sCords, ",")(5)), Val(Split(sCords, ",")(0)), Val(Split(sCords, ",")(1))
        Set PicState2 = UserControl.Image
        If Save = True Then PropertyChanged "PicState2"
        UserControl.Cls
               
        RenderButton PicState3, UserControl.hDC, Val(Split(sCords, ",")(6)), Val(Split(sCords, ",")(7)), Val(Split(sCords, ",")(0)), Val(Split(sCords, ",")(1))
        Set PicState3 = UserControl.Image
       If Save = True Then PropertyChanged "PicState3"
        UserControl.Cls
    End If
  blnBusy = False
      Draw
        
End Sub
Private Sub UserControl_Resize()
    '2nd Call
    
If SettingsLoaded = True Then
'Debug.Print "Resize " & Height
ResizenSave
Else
Draw
End If

End Sub

Private Sub UserControl_Show()
Draw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Dim nFont As StdFont
    With PropBag
        .WriteProperty "Font", demo_caption.Font, nFont
        
        .WriteProperty "theme_file_no", ntheme_file_no, 0
        .WriteProperty "theme_no", ntheme_no, 0
        .WriteProperty "Forecolor", demo_caption.ForeColor, &H0
        .WriteProperty "Backcolor", demo_caption.BackColor, &H8000000F
        .WriteProperty "Caption", demo_caption.Caption, Ambient.DisplayName
        .WriteProperty "Enabled", CLng(UserControl.Enabled), "0"
        .WriteProperty "ToolbarMode", ToolbarMode
        .WriteProperty "PicState1", PicState1, Nothing
        .WriteProperty "PicState2", PicState2, Nothing
        .WriteProperty "PicState3", PicState3, Nothing
        .WriteProperty "Skin_Cords", sCords, "0,0,0,0,0,0,0,0"
        
        If HasPicture = True Then
        .WriteProperty "BtnPicture", BtnPicture, Nothing
        Else
        .WriteProperty "BtnPicture", Nothing, Nothing
        End If
        
        Draw ' First Draw Called
    End With 'PROPBAG

End Sub

':)Code Fixer V3.0.9 (1/2/2008 10:19:14 PM) 66 + 958 = 1024 Lines Thanks Ulli for inspiration and lots of code.

