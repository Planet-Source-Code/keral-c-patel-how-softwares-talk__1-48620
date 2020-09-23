VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   4440
   ClientLeft      =   5445
   ClientTop       =   3540
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "FRMMAIN.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   4455
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   4230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   360
      TabIndex        =   5
      Top             =   4395
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2010
      TabIndex        =   7
      Top             =   165
      Width           =   330
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   2025
      Shape           =   3  'Circle
      Top             =   210
      Width           =   330
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stand By"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1170
      TabIndex        =   6
      Top             =   1395
      Width           =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   3405
      TabIndex        =   4
      Top             =   1935
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   3135
      TabIndex        =   3
      Top             =   3045
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   1995
      TabIndex        =   2
      Top             =   3660
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   705
      TabIndex        =   1
      Top             =   3045
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   1920
      Width           =   600
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   600
      Index           =   4
      Left            =   3390
      Shape           =   3  'Circle
      Top             =   1755
      Width           =   645
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   600
      Index           =   3
      Left            =   3135
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   645
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   600
      Index           =   2
      Left            =   1980
      Shape           =   3  'Circle
      Top             =   3510
      Width           =   645
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   600
      Index           =   1
      Left            =   690
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   645
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   600
      Index           =   0
      Left            =   405
      Shape           =   3  'Circle
      Top             =   1755
      Width           =   645
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private ResultRegion As Long
'For checking if the player is playing or not
Dim blnPlaying As Boolean
'For storing the filename of the song
Dim strfilename As String
'Counter
Dim i As Byte


Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)
    ObjectRegion = CreateEllipticRgn(0 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, -1 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 297 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 296 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function
Private Sub Form_Load()
        Winsock1.LocalPort = "1214"
        'Wait for Remote to Connect with our Player
        Winsock1.Listen
    
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 4
Shape1(i).FillColor = &H8000&
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
End Sub
Private Sub Label1_Click(Index As Integer)
Select Case Index
Case 0
If blnPlaying = True Then
'This is the command for pausing the MMControl.
MMControl1.Command = "Pause"
End If
Case 1
Call procPlay
Case 2
CommonDialog1.ShowOpen
Label2.Caption = CommonDialog1.FileTitle
strfilename = CommonDialog1.FileName
Case 3
MMControl1.Command = "Stop"
Case 4
frmAbout.Show
End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1(Index).FillColor = &HFF00&
End Sub

Private Sub procPlay()
'This three lines makes our MMControl ready.
MMControl1.Notify = False
MMControl1.Shareable = False
MMControl1.Wait = True
'This line assigns the filename to MMControl.
MMControl1.FileName = strfilename
'This are Commands for MMControl. I think they are self explanatory.
MMControl1.Command = "Open"
MMControl1.Command = "Prev"
MMControl1.Command = "Play"
'Set flag to true.
blnPlaying = True
End Sub

Private Sub Label3_Click()
If Winsock1.State <> sckClosed Then Winsock1.Close
Unload Me
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    
    On Error Resume Next
    'First Check if the Winsock Control is Connected or not
    'If connected then Close it

    If Winsock1.State <> sckClosed Then Winsock1.Close

    'Now accept the Request
    Winsock1.Accept requestID

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    On Error Resume Next
    Dim str As String
    'Now we will store data that has came into this string
    Winsock1.GetData str
If Mid(str, 1, 4) = "File" Then strfilename = Trim(Mid(str, 5, 1000))
If Mid(str, 1, 6) = "FTitle" Then Label2.Caption = Trim(Mid(str, 7, 255))
If LCase(str) = "play" Then Call procPlay
If LCase(str) = "pause" Then MMControl1.Command = "Pause"
If LCase(str) = "stop" Then MMControl1.Command = "Stop"
If LCase(str) = "about" Then frmAbout.Show
End Sub

