VERSION 5.00
Begin VB.UserControl SoftButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "SoftButton.ctx":0000
   Begin VB.Timer tmrMouseMonitor 
      Interval        =   10
      Left            =   1920
      Top             =   3120
   End
   Begin VB.Label Cover 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXXXX"
      ForeColor       =   &H80000010&
      Height          =   190
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
End
Attribute VB_Name = "SoftButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''border colors
Const lBorderLightStyle = &HE0E0E0
Const lBorderDarkStyle = &H808080
Const lBorderLightNormal = &HE0E0E0
Const lBorderDarkNormal = &H808080
Const lBorderOverDark = &H0
Const lBorderOverLight = &HFFFFFF

''display
Const lBorderCut = 10
Const lCaptionCut = 20

Enum BorderType
    btNormal = 0
    btLight = 1
    btDark = 2
End Enum

Dim bIsMouseDown As Boolean         ''is mouse button currently down?
Dim bMouseDownDrawn As Boolean      ''have the borders been draw to show above?
Dim bMouseOver As Boolean           ''is the mouse over the control?
Dim bIsEnabled As Boolean           ''is the control enabled
Dim btBorder As BorderType          ''the current border style

Public Event Click()
Public Event MouseOver()
Public Event MouseOut()

Private Sub tmrMouseMonitor_Timer()
    ''keeps track of the position of the mouse-cursor

    If bIsEnabled Then
        Dim udtCursorPos As POINTAPI
        Dim lhWnd As Long
        
        GetCursorPos udtCursorPos
        lhWnd = WindowFromPoint(udtCursorPos.X, _
                                udtCursorPos.Y)
        
        If lhWnd = UserControl.hWnd Then
            If Not bMouseOver Then
                bMouseOver = True
                ''
                ''mouse has just entered the control
                ''
                DrawBordersOver bIsMouseDown
            
                ''raise the mouseover event
                RaiseEvent MouseOver
            End If
        Else
            If bMouseOver Then
                bMouseOver = False
                ''
                ''mouse has just left the control
                ''
                DrawBorders False
            
                ''raise the mouseout event
                RaiseEvent MouseOut
            End If
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    bIsMouseDown = False
    bMouseOver = False
End Sub

Private Sub Cover_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If bIsEnabled And Button = 1 Then
        ''let the rest of the control know that the mouse is down
        ''and draw borders
        bIsMouseDown = True
        bMouseDownDrawn = False
        DrawBordersOver True
    End If
End Sub

Private Sub Cover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bIsMouseDown Then
        With UserControl
            If X >= 0 And X <= .Width And Y >= 0 And Y <= .Height Then
                If Not bMouseDownDrawn Then
                    DrawBordersOver True
                    bMouseDownDrawn = True
                End If
            Else
                If bMouseDownDrawn Then
                    DrawBorders False
                    bMouseDownDrawn = False
                End If
            End If
        End With
    End If
End Sub

Private Sub Cover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bIsEnabled Then
        bIsMouseDown = False
        'DrawBordersOver False
    
        If X >= 0 And X <= UserControl.Width And Y >= 0 And Y <= UserControl.Height Then
            DrawBordersOver False
            If Button = 1 Then
                ''mouse button was released over the control
                RaiseEvent Click
            End If
        Else
            DrawBorders False
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    lblCaption.Caption = UserControl.Extender.Name
    lblCaption.ForeColor = vbButtonText
    btBorder = btNormal
    bIsEnabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        lblCaption.Caption = .ReadProperty("CAPTION", UserControl.Extender.Name)
        bIsEnabled = CBool(.ReadProperty("ENABLED", "True"))
        btBorder = CInt(.ReadProperty("BORDER", 0))
        
        If bIsEnabled Then
            lblCaption.ForeColor = vbButtonText
        Else
            lblCaption.ForeColor = vbButtonShadow
        End If
        
        DrawBorders False
    End With
    DoEvents
End Sub

Private Sub UserControl_Resize()
    UserControl.Cls
    DrawBorders bIsMouseDown
    ResizeCaption
    ResizeCover
End Sub

Public Sub DrawBorders(bPressed As Boolean)
    ''controls the appearance of the control
    
    With UserControl
        
        If bPressed Then
            Line (0, 0)-(0, (.Height - lBorderCut)), lBorderDarkNormal
            Line (0, 0)-((.Width - lBorderCut), 0), lBorderDarkNormal
            
            Line (0, (.Height - lBorderCut))-((.Width - lBorderCut), (.Height - lBorderCut)), lBorderLightNormal
            Line ((.Width - lBorderCut), 0)-((.Width - lBorderCut), .Height), lBorderLightNormal ''does not use border cut
                                                                                           ''because covers up corner
                                                                                           ''space
        Else
            Select Case btBorder
                Case 0 ''normal
                    lBorderLight = lBorderLightNormal
                    lBorderDark = lBorderDarkNormal
                Case 1 ''light
                    lBorderLight = lBorderLightStyle
                    lBorderDark = lBorderLightStyle
               Case 2 ''dark
                    lBorderLight = lBorderDarkStyle
                    lBorderDark = lBorderDarkStyle
               Case Else ''fuck up
                    MsgBox "SoftTextCtl::DrawBorders()" & vbNewLine & vbNewLine & _
                           "Invalid value for 'btBorder' border type!", vbExclamation, _
                           "Ooooops!"
            End Select
            
            Line (0, 0)-(0, (.Height - lBorderCut)), lBorderLight
            Line (0, 0)-((.Width - lBorderCut), 0), lBorderLight
            
            Line (0, (.Height - lBorderCut))-((.Width - lBorderCut), (.Height - lBorderCut)), lBorderDark
            Line ((.Width - lBorderCut), 0)-((.Width - lBorderCut), .Height), lBorderDark ''does not use border cut
                                                                                          ''because covers up corner
                                                                                          ''space
        End If
        
        DoEvents
    End With
End Sub

Public Sub DrawBordersOver(bPressed As Boolean)
    ''controls the appearance of the control
    With UserControl
        
        If bPressed Then
            Line (0, 0)-(0, (.Height - lBorderCut)), lBorderOverDark
            Line (0, 0)-((.Width - lBorderCut), 0), lBorderOverDark
            
            Line (0, (.Height - lBorderCut))-((.Width - lBorderCut), (.Height - lBorderCut)), lBorderOverLight
            Line ((.Width - lBorderCut), 0)-((.Width - lBorderCut), .Height), lBorderOverLight ''does not use border cut
                                                                                               ''because covers up corner
                                                                                               ''space
        Else
            Line (0, 0)-(0, (.Height - lBorderCut)), lBorderOverLight
            Line (0, 0)-((.Width - lBorderCut), 0), lBorderOverLight
            
            Line (0, (.Height - lBorderCut))-((.Width - lBorderCut), (.Height - lBorderCut)), lBorderOverDark
            Line ((.Width - lBorderCut), 0)-((.Width - lBorderCut), .Height), lBorderOverDark ''does not use border cut
                                                                                              ''because covers up corner
                                                                                              ''space
        End If
        
        DoEvents
    End With
End Sub

Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)
    lblCaption.Caption = sNewValue
End Property

Public Sub ResizeCaption()
    With lblCaption
        .Left = lCaptionCut
        .Width = UserControl.Width - lCaptionCut
        
        .Top = CLng((UserControl.Height / 2) - (.Height / 2))
    End With
    DoEvents
End Sub

Public Sub ResizeCover()
    With Cover
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height
        DoEvents
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("CAPTION", lblCaption.Caption)
        Call .WriteProperty("ENABLED", bIsEnabled)
        Call .WriteProperty("BORDER", btBorder)
    End With
End Sub

Public Property Get Enabled() As Boolean
    Enabled = bIsEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    bIsEnabled = vNewValue
    If bIsEnabled Then
        lblCaption.ForeColor = vbButtonText
    Else
        lblCaption.ForeColor = vbButtonShadow
    End If
    DoEvents
End Property

Public Property Get BorderStyle() As BorderType
    BorderStyle = btBorder
End Property

Public Property Let BorderStyle(ByVal btNewValue As BorderType)
    btBorder = btNewValue
    DrawBorders False
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

