VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "MultiMedia Player with Remote Control.  Made By:- Keral.C.Patel."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   420
      TabIndex        =   0
      Top             =   465
      Width           =   2220
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const IDANI_OPEN = &H1
Const IDANI_CLOSE = &H2
Const IDANI_CAPTION = &H3
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawAnimatedRects Lib "User32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Sub Form_Load()
    Dim rSource As RECT, rDest As RECT, ScreenWidth As Long, ScreenHeight As Long
    'retrieve the screen width and height
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    'set the source and destination rects
    SetRect rSource, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight
    SetRect rDest, 0, 0, 200, 200
    'animate
    DrawAnimatedRects Me.hWnd, IDANI_CLOSE Or IDANI_CAPTION, rSource, rDest
    'set the form's position
    Me.Move 0, 0, 200 * Screen.TwipsPerPixelX, 200 * Screen.TwipsPerPixelY
End Sub
