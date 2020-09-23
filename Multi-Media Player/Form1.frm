VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5460
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   495
      Index           =   8
      Left            =   2625
      TabIndex        =   10
      Top             =   4575
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Learn More"
      Height          =   495
      Index           =   7
      Left            =   2625
      TabIndex        =   9
      Top             =   3990
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read-Me"
      Height          =   495
      Index           =   6
      Left            =   2640
      TabIndex        =   8
      Top             =   3420
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2250
      Top             =   5265
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   495
      Index           =   5
      Left            =   540
      TabIndex        =   7
      Top             =   4515
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   615
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   4095
      Width           =   1260
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1785
      Top             =   5265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&About"
      Height          =   495
      Index           =   4
      Left            =   540
      TabIndex        =   5
      Top             =   2835
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Stop"
      Height          =   495
      Index           =   3
      Left            =   540
      TabIndex        =   4
      Top             =   2235
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "P&ause"
      Height          =   495
      Index           =   2
      Left            =   540
      TabIndex        =   3
      Top             =   1635
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Play"
      Height          =   495
      Index           =   1
      Left            =   540
      TabIndex        =   2
      Top             =   1035
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Open"
      Height          =   495
      Index           =   0
      Left            =   540
      TabIndex        =   1
      Top             =   435
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter IP Here:-"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   660
      TabIndex        =   6
      Top             =   3810
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const BS_NULL = 1
Private mHover As Boolean
Dim InitBTStyle As Long
Private Const GWL_STYLE = (-16)
Public Sub SetInitialBTStyle(BT As CommandButton)
    
    If GetWindowLong&(BT.hwnd, GWL_STYLE) = InitBTStyle Then Exit Sub
    
    SetWindowLong& BT.hwnd, GWL_STYLE, InitBTStyle
    BT.Refresh
End Sub
Public Sub GetInitialBTStyle(BT As CommandButton)
    
    InitBTStyle = GetWindowLong&(BT.hwnd, GWL_STYLE)
End Sub
Public Sub BTFlat(BT As CommandButton)
    
    If GetWindowLong&(BT.hwnd, GWL_STYLE) And BS_FLAT Then Exit Sub
    
    SetWindowLong BT.hwnd, GWL_STYLE, InitBTStyle Or BS_FLAT
    BT.Refresh
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
CommonDialog1.ShowOpen
Winsock1.SendData "File" & CommonDialog1.FileName
Case 1
Winsock1.SendData "play"
Case 2
Winsock1.SendData "pause"
Case 3
Winsock1.SendData "stop"
Case 4
Winsock1.SendData "about"
Case 5
Winsock1.Connect Text1, "1214"
Text1.Visible = False
Label1.Visible = False
Command1(5).Top = 7000
Command1(6).Left = 540
Command1(7).Left = 540
Command1(8).Left = 540
Case 6
ShellExecute Me.hwnd, vbNullString, App.Path & "\Read-Me.htm", vbNullString, "C:\", SW_SHOWNORMAL
Case 7
ShellExecute Me.hwnd, vbNullString, "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=48418&lngWId=1", vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL
Case 8
    If Winsock1.State <> sckClosed Then Winsock1.Close
Unload Me
End Select
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If mHover Then SetInitialBTStyle Command1(Index)
    
End Sub

Private Sub Form_Load()
For i = 0 To 8
    GetInitialBTStyle Command1(i)
    mHover = True
    BTFlat Command1(i)
Next

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 8
    If mHover Then BTFlat Command1(i)
    Next
End Sub
