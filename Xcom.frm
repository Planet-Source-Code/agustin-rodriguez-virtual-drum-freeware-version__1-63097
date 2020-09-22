VERSION 5.00
Begin VB.Form Xcom 
   Caption         =   "Xcom"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2100
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   2100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Xcom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessageSTRING Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_SETTEXT As Long = &HC
Private Com_hwnd As Long

Public Event DataReceived(data As String)
Public Connected As Boolean

Public Function Start(nome As String)

    Me.Caption = nome

End Function

Public Sub Connect(nome As String)

    Com_hwnd = FindWindow(vbNullString, nome)
    Com_hwnd = FindWindowEx(Com_hwnd, ByVal 0&, vbNullString, "Text1")
    If Com_hwnd Then
        Connected = True
    End If

End Sub

Public Sub Send(x As String)

    SendMessageSTRING Com_hwnd, WM_SETTEXT, Len(x), x

End Sub

Public Sub Quit()

    Connected = False
    Text1.Text = "Text1"

End Sub

Private Sub Text1_Change()

  Dim data As String

    data = Text1.Text
    RaiseEvent DataReceived(data)

End Sub


