VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   7920
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Home Page"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Agustin Rodriguez"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3675
      Left            =   7920
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   1890
      TabIndex        =   2
      Top             =   0
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   4200
      Picture         =   "Form2.frx":16BEE
      Top             =   2880
      Width           =   3720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   6090
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9840
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Msg As String
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1


Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessageSTRING Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_SETTEXT As Long = &HC
Private Com_hWnd As Long

Private Const LWA_COLORKEY As Long = &H1
Private Const LWA_ALPHA As Long = &H2
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function apiSetWindowPos Lib "User32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1


Private Sub Form_Load()

  
  Dim Ret As Long
  Dim transColor As Long
  Dim X As String

    apiSetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
  
    transColor = RGB(255, 255, 255)

    SetLayeredWindowAttributes Me.hWnd, transColor, 50, LWA_COLORKEY
    
    Me.Caption = "Program 2"
    Start_com
  
        
  
End Sub

Private Sub Form_Paint()
Static vez As Integer
Dim X As String

    If vez = 0 Then
        vez = 1
        If Com_hWnd Then
            X = "START COMUNICATION"
            SendMessageSTRING Com_hWnd, WM_SETTEXT, Len(X), X
        End If
    End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Frame1.Visible = False Then Frame1.Visible = True
End Sub

Private Sub Label2_Click()
    Say ("Bye")
    End
End Sub

Private Sub Label3_Click()
PopupMenu Menu
End Sub

Private Sub Label4_Click(Index As Integer)


On Error Resume Next
Select Case Index
Case 1
    ShellExecute hWnd, "open", "mailto:virtual_guitar_1@hotmail.com", vbNullString, vbNullString, conSwNormal
Case 2
    ShellExecute hWnd, "open", "http://geocities.com/virtual_quality/", vbNullString, vbNullString, conSwNormal
Case 3
    Frame1.Visible = False
End Select

End Sub

Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Static ultimo As Integer
Label4(ultimo).ForeColor = 0
Label4(Index).ForeColor = &HFF
ultimo = Index

End Sub

Private Sub Text1_Change()

If Text1.Text = "Text1" Then Exit Sub
   
    Select Case Text1.Text
      Case "START COMUNICATION"
        If Not Com_hWnd Then
            Start_com
        End If
        Say ("Hi")
      Case "Hi"
        Say ("Hello! How Are You?")
      Case "Hello! How Are You?"
        Say ("I'm fine and you?")
      Case "I'm fine and you?"
        Say ("I'm happy")
      Case "I'm happy"
        Say ("Why?")
      Case "Why?"
        Say ("Because I am talking with you!")
      Case "Because I am talking with you!"
        Say ("Hi")
      Case "Bye"
        Text1.Text = "Text1" 'RESTORE THE CAPTION Text1 TO PERMIT NEW FIND WINDOW
    End Select


End Sub

Private Sub Say(X As String)
    
    Msg = X
    Timer1.Enabled = True
    
End Sub

Private Sub Start_com()

    Com_hWnd = FindWindow(vbNullString, "Program 1")
    Com_hWnd = FindWindowEx(Com_hWnd, ByVal 0&, vbNullString, "Text1")

End Sub


Private Sub Timer1_Timer()
    Label1.Caption = Msg
    If Com_hWnd Then SendMessageSTRING Com_hWnd, WM_SETTEXT, Len(Msg), Msg
    Timer1.Enabled = False
End Sub
