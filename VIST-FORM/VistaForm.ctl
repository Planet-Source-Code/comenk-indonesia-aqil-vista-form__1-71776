VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl VistaForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0F0&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
   Begin MSComctlLib.ImageList IMG_Button 
      Left            =   1080
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":0000
            Key             =   "Min_Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":053E
            Key             =   "Max_Normal"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":0A7C
            Key             =   "Cls_Normal"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":0FBA
            Key             =   "Min_Hover"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":14F8
            Key             =   "Max_Hover"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":1A36
            Key             =   "Cls_Hover"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":1F74
            Key             =   "Min_Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":24B2
            Key             =   "Max_Down"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":29F0
            Key             =   "Cls_Down"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":2F2E
            Key             =   "Min_Disable"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":346C
            Key             =   "Max_Disable"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":39AA
            Key             =   "Cls_Disable"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IMG_Disable 
      Left            =   540
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   87
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":3EE8
            Key             =   "TopLeft"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":4762
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":4874
            Key             =   "BottomLeft"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":4986
            Key             =   "Top"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":51B8
            Key             =   "TopRight"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":5A32
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":5B2C
            Key             =   "BottomRight"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":5C3E
            Key             =   "Bottom"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IMG_Active 
      Left            =   0
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   120
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":5D70
            Key             =   "TopLeft"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":6902
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":69E4
            Key             =   "BottomLeft"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":6AF6
            Key             =   "Top"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":6DE8
            Key             =   "TopRight"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":797A
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":7A8C
            Key             =   "BottomRight"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VistaForm.ctx":7B9E
            Key             =   "Bottom"
         EndProperty
      EndProperty
   End
   Begin VB.Image iCapt 
      Height          =   135
      Left            =   90
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image iConX 
      Height          =   240
      Left            =   120
      Picture         =   "VistaForm.ctx":7CB0
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image szBottom 
      Height          =   120
      Left            =   990
      Top             =   360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image szCorner 
      Height          =   120
      Left            =   1440
      Top             =   360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image szRight 
      Height          =   240
      Left            =   1440
      Top             =   90
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image iCls 
      Height          =   225
      Left            =   990
      Top             =   90
      Width           =   420
   End
   Begin VB.Image iMax 
      Height          =   225
      Left            =   540
      Top             =   90
      Width           =   420
   End
   Begin VB.Image iMin 
      Height          =   225
      Left            =   90
      Top             =   90
      Width           =   420
   End
   Begin VB.Image iTitleBar 
      Height          =   135
      Left            =   270
      Top             =   360
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "VistaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'=====================================Vista-Form v.1=============================================
' VistaForm ActiveX Control (written on Feb-13-09)
' Copyright © 2009 by Wongjava Software. All rights reserved.
' Author    : Sudarmono-[Comenk]
' E-Mail    : wongjava_bsi@yahoo.com
' Notes     : If anyone have a problem when using this ActiveX please mail me.
'             If anyone have an idea to completed this ActiveX please mail me.
'             and I'am sorry for my English...
'
' Not Complete Yet  : 1. When Maximize I can't show the taskbar.
'                     2. I Can't make simple hook for known Parent Form is Active/UnActive
'================================================================================================
'\\API DECLARATION
'================================================================================================
Private Declare Function SetRect Lib "User32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, _
                ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "User32.dll" (ByRef lpRect As RECT, _
                ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "User32.dll" (ByRef lpDestRect As RECT, _
                ByRef lpSourceRect As RECT) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
                ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, _
                ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, _
                ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, _
                ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" _
                (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "User32.dll" Alias "DrawTextA" (ByVal hDC As Long, _
                ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, _
                ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "User32.dll" (ByVal hDC As Long, ByVal lpStr As Long, _
                ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
'================================================================================================
'\\TYPE, CONST AND ENUM DECLARATION
'================================================================================================
Private Const DT_CALCRECT               As Long = &H400
Private Const DT_CENTER                 As Long = &H1
Private Const DT_NOCLIP                 As Long = &H100
Private Const DT_WORDBREAK              As Long = &H10
Private Const DT_CALCFLAG               As Long = DT_WORDBREAK Or DT_CALCRECT Or DT_NOCLIP Or DT_CENTER
Private Const DT_DRAWFLAG               As Long = DT_WORDBREAK Or DT_NOCLIP Or DT_CENTER
Private Const VER_PLATFORM_WIN32_NT     As Long = 2
Private Const RGN_OR                    As Long = 2
Private Const RGN_AND                   As Long = 1
Private Const RGN_XOR                   As Long = 3
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Enum aqUC_State
    stNormal = 0
    stHover = 1
    stDown = 2
    stDisable = 3
End Enum

Enum aqUC_BldStyle
    bsFixedSingle = 0
    bsSizeable = 1
End Enum

Enum aqUC_WinState
    wsNormal = 0
    wsMinimized = 1
    wsMaximized = 2
End Enum

Private Type UC_Setting
    FirstInit   As Boolean
    CaptRect    As RECT
    PrntActive  As Boolean
    State       As aqUC_WinState
    WiNT        As Boolean
End Type

Private Type UC_Properties
    BackColor   As OLE_COLOR
    BolderStyle As aqUC_BldStyle
    Caption     As String
    ControlBox  As Boolean
    Font        As StdFont
    ForeColor   As OLE_COLOR
    Icon        As StdPicture
    MDIChild    As Boolean
    MinButton   As Boolean
    MaxButton   As Boolean
    Moveable    As Boolean
    ShowInTask  As Boolean
    WinState    As aqUC_WinState
End Type

Private Type UC_FormPictButton
    Min_Normal  As StdPicture
    Min_Hover   As StdPicture
    Min_Down    As StdPicture
    Min_Disable As StdPicture
    Max_Normal  As StdPicture
    Max_Hover   As StdPicture
    Max_Down    As StdPicture
    Max_Disable As StdPicture
    Cls_Normal  As StdPicture
    Cls_Hover   As StdPicture
    Cls_Down    As StdPicture
    Cls_Disable As StdPicture
End Type

Private Type UC_FormPictActive
    TopLeft     As StdPicture
    Top         As StdPicture
    TopRight    As StdPicture
    Left        As StdPicture
    Right       As StdPicture
    BottomLeft  As StdPicture
    Bottom      As StdPicture
    BottomRight As StdPicture
End Type

Private Type UC_FormPictDisable
    TopLeft     As StdPicture
    Top         As StdPicture
    TopRight    As StdPicture
    Left        As StdPicture
    Right       As StdPicture
    BottomLeft  As StdPicture
    Bottom      As StdPicture
    BottomRight As StdPicture
End Type
'================================================================================================
'\\LOCAL DECLARATION
'================================================================================================
Private m_pictbutton    As UC_FormPictButton
Private m_pictactive    As UC_FormPictActive
Private m_pictdisable   As UC_FormPictDisable
Private m_setting       As UC_Setting
Private m_property      As UC_Properties

Private Sub iCls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.ControlBox) Or (Not m_setting.PrntActive) Then Exit Sub
    Call DrawUC_Button(True, 2, stDown)
End Sub

Private Sub iCls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.ControlBox) Or (Not m_setting.PrntActive) Then Exit Sub
    If (m_property.MinButton) Then Call DrawUC_Button(True, 0, stNormal)
    If (m_property.MaxButton) Then Call DrawUC_Button(True, 1, stNormal)
    Call DrawUC_Button(True, 2, stHover)
End Sub

Private Sub iCls_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.ControlBox) Or (Not m_setting.PrntActive) Then Exit Sub
    
    Unload UserControl.Parent
End Sub

Private Sub iMax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.MaxButton) Or (Not m_setting.PrntActive) Then Exit Sub
    Call DrawUC_Button(True, 1, stDown)
End Sub

Private Sub iMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.MaxButton) Or (Not m_setting.PrntActive) Then Exit Sub
    If (m_property.MinButton) Then Call DrawUC_Button(True, 0, stNormal)
    If (m_property.ControlBox) Then Call DrawUC_Button(True, 2, stNormal)
    Call DrawUC_Button(True, 1, stHover)
End Sub

Private Sub iMax_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.MaxButton) Or (Not m_setting.PrntActive) Then Exit Sub
    
    If (Button = 1) Then
        If (Not m_setting.State = wsMaximized) Then
            UserControl.Parent.WindowState = vbMaximized
            m_setting.State = wsMaximized
        Else
            UserControl.Parent.WindowState = vbNormal
            m_setting.State = wsNormal
        End If
        Call DrawUC_Active
    End If
End Sub

Private Sub iMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.MinButton) Or (Not m_setting.PrntActive) Then Exit Sub
    Call DrawUC_Button(True, 0, stDown)
End Sub

Private Sub iMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.MinButton) Or (Not m_setting.PrntActive) Then Exit Sub
    If (m_property.MaxButton) Then Call DrawUC_Button(True, 1, stNormal)
    If (m_property.ControlBox) Then Call DrawUC_Button(True, 2, stNormal)
    Call DrawUC_Button(True, 0, stHover)
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_property.BackColor
End Property
Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    m_property.BackColor = nValue
    
    PropertyChanged "BackColor"
End Property
Public Property Get BorderStyle() As aqUC_BldStyle
    BorderStyle = m_property.BolderStyle
End Property
Public Property Let BorderStyle(ByVal nValue As aqUC_BldStyle)
    m_property.BolderStyle = nValue
    
    If (m_property.BolderStyle = bsFixedSingle) Then
        m_property.MinButton = False
        m_property.MaxButton = False
    Else
        m_property.MinButton = True
        m_property.MaxButton = True
    End If
    
    If (Not Ambient.UserMode) Then
        Call DrawUC_IDE
    End If
    
    PropertyChanged "MinButton"
    PropertyChanged "MaxButton"
    PropertyChanged "BorderStyle"
End Property
Public Property Get Caption() As String
    Caption = m_property.Caption
End Property
Public Property Let Caption(ByVal nValue As String)
    m_property.Caption = nValue
    
    If (Not Ambient.UserMode) Then DrawUC_IDE
    UserControl.Parent.Caption = nValue
    
    PropertyChanged "Caption"
End Property
Public Property Get ControlBox() As Boolean
    ControlBox = m_property.ControlBox
End Property
Public Property Let ControlBox(ByVal nValue As Boolean)
    m_property.ControlBox = nValue
    
    PropertyChanged "ControlBox"
End Property
Public Property Get Font() As StdFont
    Set Font = m_property.Font
End Property
Public Property Set Font(ByVal nValue As StdFont)
    Set m_property.Font = nValue
    Set UserControl.Font = nValue
    If (Not Ambient.UserMode) Then
        DrawUC_IDE
    End If
    PropertyChanged "Font"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_property.ForeColor
End Property
Public Property Let ForeColor(ByVal nValue As OLE_COLOR)
    m_property.ForeColor = nValue
    If (Not Ambient.UserMode) Then
        DrawUC_IDE
    End If
    PropertyChanged "ForeColor"
End Property
Public Property Get Icon() As StdPicture
    Set Icon = m_property.Icon
End Property
Public Property Set Icon(ByVal nValue As StdPicture)
    Set m_property.Icon = nValue
    
    Set iConX.Picture = nValue
    If (Not Ambient.UserMode) Then
        DrawUC_IDE
    End If
    
    PropertyChanged "Icon"
End Property
Public Property Get MaxButton() As Boolean
    MaxButton = m_property.MaxButton
End Property
Public Property Let MaxButton(ByVal nValue As Boolean)
    m_property.MaxButton = nValue
    
    If (Not m_property.MaxButton) And (Not m_property.MinButton) Then
        m_property.BolderStyle = bsFixedSingle
    Else
        m_property.BolderStyle = bsSizeable
    End If
    
    If (Not Ambient.UserMode) Then
        Call DrawUC_IDE
    End If
    
    PropertyChanged "BorderStyle"
    PropertyChanged "MaxButton"
End Property
Public Property Get MinButton() As Boolean
    MinButton = m_property.MinButton
End Property
Public Property Let MinButton(ByVal nValue As Boolean)
    m_property.MinButton = nValue
    
    If (Not m_property.MaxButton) And (Not m_property.MinButton) Then
        m_property.BolderStyle = bsFixedSingle
    Else
        m_property.BolderStyle = bsSizeable
    End If
    
    If (Not Ambient.UserMode) Then
        Call DrawUC_IDE
    End If
    
    PropertyChanged "BorderStyle"
    PropertyChanged "MinButton"
End Property
Public Property Get MDIChild() As Boolean
    MDIChild = m_property.MDIChild
End Property
Public Property Get Moveable() As Boolean
    Moveable = m_property.Moveable
End Property
Public Property Let Moveable(ByVal nValue As Boolean)
    m_property.Moveable = nValue
    
    PropertyChanged "Moveable"
End Property
Public Property Get WinState() As aqUC_WinState
    WinState = m_property.WinState
End Property
Public Property Let WinState(ByVal nValue As aqUC_WinState)
    m_property.WinState = nValue
    
    PropertyChanged "WinState"
End Property
Friend Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
Friend Property Get hDC() As Long
    hDC = UserControl.hDC
End Property
Friend Function ReFresh()
    UserControl.ReFresh
    UserControl.Parent.ReFresh
End Function
Friend Property Get ParenthWnd() As Long
    ParenthWnd = UserControl.Parent.hWnd
End Property
Private Sub iMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.MinButton) Or (Not m_setting.PrntActive) Then Exit Sub
    
    If (Button = 1) Then
        If (Not m_setting.State = wsMinimized) Then
            UserControl.Parent.WindowState = vbMinimized
            m_setting.State = wsMinimized
        Else
            UserControl.Parent.WindowState = vbNormal
            m_setting.State = wsNormal
        End If
        Call DrawUC_Active
    End If
End Sub

Private Sub iTitleBar_DblClick()
    If (m_property.MaxButton) Then
        Call iMax_MouseUp(vbLeftButton, 0, 0, 0)
    End If
End Sub

Private Sub iTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (m_setting.State = wsMaximized) Then Exit Sub
    If (Not m_property.Moveable) Then Exit Sub
    
    If (Button = 1) Then
        Call ReleaseCapture
        Call SendMessage(Me.ParenthWnd, WM_NCLBUTTONDOWN, HTCAPTION, &O0)
    End If
End Sub

Private Sub iTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub szBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.BolderStyle = bsSizeable) Or (Not m_setting.PrntActive) Then Exit Sub
    
    szBottom.MousePointer = 7
    
    If (Button = 1) Then
        With UserControl.Parent
            If (.Height + Y) > 525 Then
                .Height = .Height + Y
            End If
        End With
        
        Call DrawUC_Active
    End If
End Sub
Private Sub szCorner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.BolderStyle = bsSizeable) Or (Not m_setting.PrntActive) Then Exit Sub
    
    szCorner.MousePointer = 8
    
    If (Button = 1) Then
        With UserControl.Parent
            
            If (.Height + Y) > 525 Then
                .Height = .Height + Y
            End If
            
            If (.Width + X) > 1830 + (TextWidth(m_property.Caption) * 15) Then
                .Width = .Width + X
            End If
        End With
        
        Call DrawUC_Active
    End If
End Sub

Private Sub szRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not m_property.BolderStyle = bsSizeable) Or (Not m_setting.PrntActive) Then Exit Sub
    
    szRight.MousePointer = 9
    If (Button = 1) Then
        With UserControl.Parent
            If (.Width + X) > 1830 + (TextWidth(m_property.Caption) * 15) Then
                .Width = .Width + X
            End If
        End With
        
        Call DrawUC_Active
    End If
End Sub

Private Sub UserControl_InitProperties()
    With m_property
        .BackColor = &HF0F0F0
        .BolderStyle = bsSizeable
        .Caption = Ambient.DisplayName
        .ControlBox = UserControl.Parent.ControlBox
    Set .Font = Ambient.Font
        .ForeColor = Ambient.ForeColor
    Set .Icon = Nothing
        .MaxButton = UserControl.Parent.MaxButton
        .MDIChild = UserControl.Parent.MDIChild
        .MinButton = UserControl.Parent.MinButton
        .Moveable = UserControl.Parent.Moveable
        .WinState = wsNormal
    End With
    
    m_setting.FirstInit = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (m_property.MinButton) Then Call DrawUC_Button(True, 0, stNormal)
    If (m_property.MaxButton) Then Call DrawUC_Button(True, 1, stNormal)
    If (m_property.ControlBox) Then Call DrawUC_Button(True, 2, stNormal)
End Sub
Private Sub UserControl_Initialize()
    m_setting.WiNT = IsPlatformNT()
    With m_pictbutton
        Set .Cls_Disable = IMG_Button.ListImages("Cls_Disable").Picture
        Set .Cls_Down = IMG_Button.ListImages("Cls_Down").Picture
        Set .Cls_Hover = IMG_Button.ListImages("Cls_Hover").Picture
        Set .Cls_Normal = IMG_Button.ListImages("Cls_Normal").Picture
        Set .Max_Disable = IMG_Button.ListImages("Max_Disable").Picture
        Set .Max_Down = IMG_Button.ListImages("Max_Down").Picture
        Set .Max_Hover = IMG_Button.ListImages("Max_Hover").Picture
        Set .Max_Normal = IMG_Button.ListImages("Max_Normal").Picture
        Set .Min_Disable = IMG_Button.ListImages("Min_Disable").Picture
        Set .Min_Down = IMG_Button.ListImages("Min_Down").Picture
        Set .Min_Hover = IMG_Button.ListImages("Min_Hover").Picture
        Set .Min_Normal = IMG_Button.ListImages("Min_Normal").Picture
    End With
    
    With m_pictactive
        Set .Bottom = IMG_Active.ListImages("Bottom").Picture
        Set .BottomLeft = IMG_Active.ListImages("BottomLeft").Picture
        Set .BottomRight = IMG_Active.ListImages("BottomRight").Picture
        Set .Left = IMG_Active.ListImages("Left").Picture
        Set .Right = IMG_Active.ListImages("Right").Picture
        Set .Top = IMG_Active.ListImages("Top").Picture
        Set .TopLeft = IMG_Active.ListImages("TopLeft").Picture
        Set .TopRight = IMG_Active.ListImages("TopRight").Picture
    End With
    
    With m_pictdisable
        Set .Bottom = IMG_Disable.ListImages("Bottom").Picture
        Set .BottomLeft = IMG_Disable.ListImages("BottomLeft").Picture
        Set .BottomRight = IMG_Disable.ListImages("BottomRight").Picture
        Set .Left = IMG_Disable.ListImages("Left").Picture
        Set .Right = IMG_Disable.ListImages("Right").Picture
        Set .Top = IMG_Disable.ListImages("Top").Picture
        Set .TopLeft = IMG_Disable.ListImages("TopLeft").Picture
        Set .TopRight = IMG_Disable.ListImages("TopRight").Picture
    End With
End Sub
Private Sub DrawUC_Caption(ByVal Color As Long, Optional pShadow As Boolean = False)
    Dim tx      As String
    Dim tn      As Long
    Dim rc      As RECT
    
    tx = m_property.Caption
    tn = Len(tx)
    
    If (tn = 0) Then Exit Sub
    
    Call SetRect(rc, 0, 0, TextWidth(m_property.Caption), TextHeight(m_property.Caption))
    Call OffsetRect(rc, IIf((Not m_property.Icon Is Nothing), iConX.Left + iConX.Width + 5, 9), ((28 - rc.Bottom) / 2) + 3)
    Call CopyRect(m_setting.CaptRect, rc)
    
    If (pShadow) Then
        Call SetTextColor(Me.hDC, vbGrayed)
        
        If (m_setting.WiNT) Then
            Call DrawTextW(Me.hDC, StrPtr(tx), tn, m_setting.CaptRect, DT_DRAWFLAG)
        Else
            Call DrawText(Me.hDC, tx, tn, m_setting.CaptRect, DT_DRAWFLAG)
        End If
        
        Call OffsetRect(rc, -1, -1)
        Call CopyRect(m_setting.CaptRect, rc)
        
        Call SetTextColor(Me.hDC, Color)
        
        If (m_setting.WiNT) Then
            Call DrawTextW(Me.hDC, StrPtr(tx), tn, m_setting.CaptRect, DT_DRAWFLAG)
        Else
            Call DrawText(Me.hDC, tx, tn, m_setting.CaptRect, DT_DRAWFLAG)
        End If
    Else
        Call OffsetRect(rc, -1, -1)
        Call CopyRect(m_setting.CaptRect, rc)
        
        Call SetTextColor(Me.hDC, Color)
        
        If (m_setting.WiNT) Then
            Call DrawTextW(Me.hDC, StrPtr(tx), tn, m_setting.CaptRect, DT_DRAWFLAG)
        Else
            Call DrawText(Me.hDC, tx, tn, m_setting.CaptRect, DT_DRAWFLAG)
        End If
    End If
End Sub
Private Sub DrawUC_Region(Optional EllipseW As Integer = 9, Optional EllipseH As Integer = 9)
    Dim hRgn    As Long
    Dim hRgn2   As Long
    Dim hRgn3   As Long
    
    If (m_setting.State = wsNormal) Or (m_setting.State = wsMaximized) Then
        With UserControl
            hRgn = CreateRoundRectRgn(0, 0, .ScaleWidth + 1, .ScaleHeight, EllipseW, EllipseH)
            hRgn2 = CreateRectRgn(0, 9, .ScaleWidth, .ScaleHeight)
            hRgn3 = CreateRectRgn(8, 28, .ScaleWidth - 8, .ScaleHeight - 8)
        End With
        
        Call CombineRgn(hRgn, hRgn, hRgn2, RGN_OR)
        Call DeleteObject(hRgn2)
        
        Call CombineRgn(hRgn, hRgn, hRgn3, RGN_XOR)
        Call DeleteObject(hRgn3)
    End If
    
    Call SetWindowRgn(Me.hWnd, hRgn, True)
    Call DeleteObject(hRgn)
    
    Call DrawPR_Region
End Sub
Private Sub DrawPR_Region(Optional EllipseW As Integer = 9, Optional EllipseH As Integer = 9)
    Dim hRgn    As Long
    Dim hRgn2   As Long
    
    If (m_setting.State = wsNormal) Then
        With UserControl.Parent
            hRgn = CreateRoundRectRgn(0, 0, .ScaleX(.Width, 1, 3) + 1, .ScaleY(.Height, 1, 3), EllipseW, EllipseH)
            hRgn2 = CreateRectRgn(0, 9, ScaleX(.Width, 1, 3), .ScaleY(.Height, 1, 3))
        End With
        
        Call CombineRgn(hRgn, hRgn, hRgn2, RGN_OR)
        Call DeleteObject(hRgn2)
    ElseIf (m_setting.State = wsNormal) Then
        With UserControl.Parent
            hRgn = CreateRectRgn(0, 0, ScaleX(.Width, 1, 3), .ScaleY(.Height, 1, 3))
        End With
    End If
    Call SetWindowRgn(Me.ParenthWnd, hRgn, True)
    Call DeleteObject(hRgn)
End Sub
Private Function IsPlatformNT() As Boolean
    Dim OSINFO As OSVERSIONINFO
        OSINFO.dwOSVersionInfoSize = Len(OSINFO)
        
    If (GetVersionEx(OSINFO)) Then
        IsPlatformNT = (OSINFO.dwPlatformId = VER_PLATFORM_WIN32_NT)
    End If
End Function
Private Sub DrawUC_Active(Optional pState As aqUC_State = stNormal, Optional pForceBtn As Boolean = True)
    With UserControl
        .Height = IIf((m_setting.State = wsNormal) Or (m_setting.State = wsMaximized), .Parent.Height, 0) 'IIf((m_setting.State = wsNormal), .Parent.Height, IIf((m_setting.State = wsMaximized), 420, 0))
        .Width = .Parent.Width
        
        .Cls
        If (m_setting.State = wsNormal) Or (m_setting.State = wsMaximized) Then
            If (.ScaleHeight > 38) Then
                .PaintPicture m_pictactive.Left, 0, 30, , .ScaleHeight - 38
                .PaintPicture m_pictactive.Right, .ScaleWidth - 8, 30, , .ScaleHeight - 38
            End If
            .PaintPicture m_pictactive.TopLeft, 0, 0
            .PaintPicture m_pictactive.TopRight, .ScaleWidth - 8, 0
            .PaintPicture m_pictactive.Top, 8, 0, .ScaleWidth - 16
            .PaintPicture m_pictactive.BottomLeft, 0, .ScaleHeight - 8
            .PaintPicture m_pictactive.BottomRight, .ScaleWidth - 8, .ScaleHeight - 8
            .PaintPicture m_pictactive.Bottom, 8, .ScaleHeight - 8, .ScaleWidth - 16
            If (Not m_property.Icon Is Nothing) Then .PaintPicture m_property.Icon, 8, 8: iConX.Left = 8
        End If
    End With
    Call DrawUC_Caption(m_property.ForeColor)
    Call DrawUC_Button
    Me.ReFresh
    Call DrawUC_Region
End Sub
Private Sub DrawUC_Disable(Optional pState As aqUC_State = stDisable)
    With UserControl
        .Cls
        .PaintPicture m_pictdisable.Left, 0, 30, , .ScaleHeight - 38
        .PaintPicture m_pictdisable.Right, .ScaleWidth - 8, 30, , .ScaleHeight - 38
        .PaintPicture m_pictdisable.TopLeft, 0, 0
        .PaintPicture m_pictdisable.TopRight, .ScaleWidth - 8, 0
        .PaintPicture m_pictdisable.Top, 8, 0, .ScaleWidth - 16
        .PaintPicture m_pictdisable.BottomLeft, 0, .ScaleHeight - 8
        .PaintPicture m_pictdisable.BottomRight, .ScaleWidth - 8, .ScaleHeight - 8
        .PaintPicture m_pictdisable.Bottom, 8, .ScaleHeight - 8, .ScaleWidth - 16
    End With
    Call DrawUC_Button(True, 0, stDisable)
    Me.ReFresh
    Call DrawUC_Region
End Sub
Private Sub DrawUC_IDE()
    With UserControl
        .Cls
        .PaintPicture m_pictdisable.TopLeft, 0, 0
        .PaintPicture m_pictdisable.TopRight, .ScaleWidth - 8, 0
        .PaintPicture m_pictdisable.Top, 8, 0, .ScaleWidth - 16
        If (Not m_property.Icon Is Nothing) Then .PaintPicture m_property.Icon, 8, 8
    End With
    Call DrawUC_Caption(m_property.ForeColor)
    Call DrawUC_Button
End Sub
Private Sub DrawUC_Button(Optional pAmbient As Boolean = False, Optional pBtn As Integer = 0, _
Optional pBtnState As aqUC_State = stNormal)
    Select Case pAmbient
        Case True
            With UserControl
                If (m_property.ControlBox) Then
                    If (pBtnState = stDisable) And (Not m_setting.PrntActive) Then
                        If (m_property.MinButton) Then .PaintPicture m_pictbutton.Min_Disable, iMin.Left, iMin.Top
                        If (m_property.MaxButton) Then .PaintPicture m_pictbutton.Max_Disable, iMax.Left, iMax.Top
                        .PaintPicture m_pictbutton.Cls_Disable, iCls.Left, iCls.Top
                        
                    ElseIf (pBtnState = stDisable) And (m_setting.PrntActive) Then
                        If (pBtn = 0) Then '\\Min
                            .PaintPicture m_pictbutton.Min_Disable, iMin.Left, iMin.Top
                        ElseIf (pBtn = 1) Then '\\Max
                            .PaintPicture m_pictbutton.Max_Disable, iMax.Left, iMax.Top
                        ElseIf (pBtn = 2) Then '\\Cls
                            .PaintPicture m_pictbutton.Cls_Disable, iCls.Left, iCls.Top
                        End If
                        
                    ElseIf (pBtnState = stDown) And (m_setting.PrntActive) Then
                        If (pBtn = 0) Then '\\Min
                            .PaintPicture m_pictbutton.Min_Down, iMin.Left, iMin.Top
                        ElseIf (pBtn = 1) Then '\\Max
                            .PaintPicture m_pictbutton.Max_Down, iMax.Left, iMax.Top
                        ElseIf (pBtn = 2) Then '\\Cls
                            .PaintPicture m_pictbutton.Cls_Down, iCls.Left, iCls.Top
                        End If
                        
                    ElseIf (pBtnState = stHover) And (m_setting.PrntActive) Then
                        If (pBtn = 0) Then '\\Min
                            .PaintPicture m_pictbutton.Min_Hover, iMin.Left, iMin.Top
                        ElseIf (pBtn = 1) Then '\\Max
                            .PaintPicture m_pictbutton.Max_Hover, iMax.Left, iMax.Top
                        ElseIf (pBtn = 2) Then '\\Cls
                            .PaintPicture m_pictbutton.Cls_Hover, iCls.Left, iCls.Top
                        End If
                        
                    ElseIf (pBtnState = stNormal) And (m_setting.PrntActive) Then
                        If (pBtn = 0) Then '\\Min
                            .PaintPicture m_pictbutton.Min_Normal, iMin.Left, iMin.Top
                        ElseIf (pBtn = 1) Then '\\Max
                            .PaintPicture m_pictbutton.Max_Normal, iMax.Left, iMax.Top
                        ElseIf (pBtn = 2) Then '\\Cls
                            .PaintPicture m_pictbutton.Cls_Normal, iCls.Left, iCls.Top
                        End If
                    End If
                End If
            End With
            
        Case False
            With UserControl
                If (m_property.ControlBox) Then
                    If (m_property.BolderStyle = bsSizeable) Then
                        If (m_property.MinButton) Then .PaintPicture m_pictbutton.Min_Normal, iMin.Left, iMin.Top _
                        Else .PaintPicture m_pictbutton.Min_Disable, iMin.Left, iMin.Top
                        
                        If (m_property.MaxButton) Then .PaintPicture m_pictbutton.Max_Normal, iMax.Left, iMax.Top _
                        Else .PaintPicture m_pictbutton.Max_Disable, iMax.Left, iMax.Top
                    End If
                    
                    .PaintPicture m_pictbutton.Cls_Normal, iCls.Left, iCls.Top
                End If
            End With
            
    End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_property.BackColor = .ReadProperty("BackColor", &HF0F0F0)
        m_property.BolderStyle = .ReadProperty("BorderStyle", bsSizeable)
        m_property.Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_property.ControlBox = .ReadProperty("ControlBox", True)
    Set m_property.Font = .ReadProperty("Font", Ambient.Font)
        m_property.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
    Set m_property.Icon = .ReadProperty("Icon", Nothing)
        m_property.MaxButton = .ReadProperty("MaxButton", True)
        m_property.MDIChild = .ReadProperty("MDIChild", False)
        m_property.MinButton = .ReadProperty("MinButton", True)
        m_property.Moveable = .ReadProperty("Moveable", True)
        m_property.WinState = .ReadProperty("WinState", wsNormal)
    End With
    
    With UserControl
    Set .Font = m_property.Font
        .ForeColor = m_property.ForeColor
        .Parent.Caption = m_property.Caption
    End With
    
    iMin.Top = 9
    iMax.Top = 9
    iCls.Top = 9
    If (Not m_property.Icon Is Nothing) Then Set iConX.Picture = m_property.Icon
    iMin.Visible = IIf((m_property.ControlBox), IIf(((m_property.MinButton) Or (m_property.MaxButton)), True, False), False)
    iMax.Visible = IIf((m_property.ControlBox), IIf(((m_property.MinButton) Or (m_property.MaxButton)), True, False), False)
    iCls.Visible = IIf((m_property.ControlBox), True, False)
    
    If (Not Ambient.UserMode) Then
        DrawUC_IDE
    End If
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", m_property.BackColor, &HF0F0F0
        .WriteProperty "BorderStyle", m_property.BolderStyle, bsSizeable
        .WriteProperty "Caption", m_property.Caption, Ambient.DisplayName
        .WriteProperty "ControlBox", m_property.ControlBox, True
        .WriteProperty "Font", m_property.Font, Ambient.Font
        .WriteProperty "ForeColor", m_property.ForeColor, Ambient.ForeColor
        .WriteProperty "Icon", m_property.Icon, Nothing
        .WriteProperty "MaxButton", m_property.MaxButton, True
        .WriteProperty "MDIChild", m_property.MDIChild, False
        .WriteProperty "MinButton", m_property.MinButton, True
        .WriteProperty "Moveable", m_property.Moveable, True
        .WriteProperty "WinState", m_property.WinState, wsNormal
    End With
End Sub
Private Sub UserControl_Resize()
    If (Not m_setting.State = wsMinimized) Then
        iMin.Left = UserControl.ScaleWidth - 95
        iMax.Left = iMin.Left + 29
        iCls.Left = iMax.Left + 29
        
        szBottom.Visible = False
        szCorner.Visible = False
        szRight.Visible = False
        iTitleBar.Visible = False
        
        If (Not Ambient.UserMode) Then
            If (UserControl.Height > 420) Or (UserControl.Height < 420) Then UserControl.Height = 420
            DrawUC_IDE
        Else
            szBottom.Move 0, UserControl.ScaleHeight - 8, UserControl.ScaleWidth - 8
            szCorner.Move UserControl.ScaleWidth - 8, UserControl.ScaleHeight - 8
            szRight.Move UserControl.ScaleWidth - 8, 8, 8, UserControl.ScaleHeight - 16
            iTitleBar.Move 0, 0, UserControl.ScaleWidth, 28
            iTitleBar.Visible = True
            
            If (m_setting.State = wsNormal) Then
                szBottom.Visible = True
                szCorner.Visible = True
                szRight.Visible = True
            End If
        End If
    Else
        If (UserControl.Parent.WindowState = 0) Then
            m_setting.State = wsNormal
            Call DrawUC_Active
        ElseIf (UserControl.Parent.WindowState = 2) Then
            m_setting.State = wsMaximized
            Call DrawUC_Active
        End If
    End If
End Sub
Private Sub UserControl_Show()
    If (Not UserControl.Parent.BorderStyle = 0) Then UserControl.Parent.BorderStyle = 0
    
    If (Ambient.UserMode) Then
        m_setting.PrntActive = True
        UserControl.Height = UserControl.Parent.ScaleHeight
        Call DrawUC_Active
    Else
        With UserControl
            If (m_setting.FirstInit) Then .Parent.BackColor = &HF0F0F0
        End With
    End If
End Sub


