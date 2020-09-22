VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3570
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   0
      TabIndex        =   2
      Top             =   -120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   435
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Run in XP only
'Excuse me for my bad English

Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessageSTRING Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_SETTEXT As Long = &HC

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

Private Com_hWnd As Long
Private Msg As String


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
    
    
    'The form must have this name
    Me.Caption = "Program 1"
    Start_com
        
End Sub

Private Sub Form_Paint()
'I use the first paint to start the communication
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

Private Sub Label2_Click()

    Say ("Bye")
    End

End Sub

Private Sub Text1_Change()

'When SendMessageSTRING sends a string to Text1 Hwnd this Event is activated.
'You can use Winproc to make this, but case to exist any bug in the program or the
'program is stopped without to restore the Oldproc, the program will freeze

    If Text1.Text = "Text1" Then
        Exit Sub
    End If

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
'I used a Timer to send message to prevent Out Space on stack and to make a interval between the messages
    Msg = X
    Timer1.Enabled = True
        
End Sub

Private Sub Start_com()
'Here I find the other Window and the Box Text
    Com_hWnd = FindWindow(vbNullString, "Program 2")
    Com_hWnd = FindWindowEx(Com_hWnd, ByVal 0&, vbNullString, "Text1")

End Sub

Private Sub Timer1_Timer()

    Label1.Caption = Msg
    If Com_hWnd Then
        SendMessageSTRING Com_hWnd, WM_SETTEXT, Len(Msg), Msg
    End If
    Timer1.Enabled = False
    
End Sub


