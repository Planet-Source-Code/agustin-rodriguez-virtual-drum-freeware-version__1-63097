VERSION 5.00
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   FillColor       =   &H00FFFFFE&
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7590
      Left            =   5895
      ScaleHeight     =   506
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   491
      TabIndex        =   0
      Top             =   6315
      Visible         =   0   'False
      Width           =   7365
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Quant_UNDO As Integer = 127

Private Passo As Double
Private Setor As Integer
Private Escolhidos As String * 16
Private Backup(256, 127) As Long
Private Backup_34ou44(127) As Integer

Private Nivel_para_Undo As Integer

Private Const MK_LBUTTON As Long = &H1
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
    
Private Const MK_RBUTTON As Long = &H2
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205

Private Const WM_LBUTTONDBLCLK As Long = &H203
    
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub Form_Load()

  Dim NormalWindowStyle As Long
  Dim ret As Long, xx As Integer, n As String
  Dim col As Long
  Dim i As Integer
  
  Static Raio As Integer

    Escolhidos = "----------------"

    Raio = 208
    FillStyle = 0
    ForeColor = &HA5B423
    Passo = 5.625 * 4
    Circle (CentroX, CentroY), Raio
    For i = 0 To 16
        ang = ang + Passo

        ag = Radiano * ang
        xt = (Sin(ag)) * Raio
        yt = -Cos(ag) * Raio
        Line (CentroX + xt - Raio * Sin(ag), CentroY + yt - Raio * -Cos(ag))-(CentroX + xt, CentroY + yt)
    Next i
    
    NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 50, LWA_ALPHA

    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    col = 255 'RGB(255, 255, 255)
    SetLayeredWindowAttributes Me.hwnd, 255, 50, LWA_COLORKEY Or LWA_ALPHA

    Prepare_Mascara
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim nMousePosition As Long
  Dim mbrush As Long
  Dim S As Long
  
    If Shift Then
        If Shift = 4 And Button = 1 Then
            Undo
            Exit Sub
        End If
        
        If Shift = 4 And Button = 2 Then
            Redo
            Exit Sub
        End If
        
        S = Picture1.Point(x, y) / 10000
        If S > 15 Then Exit Sub
        Setor = S
        
        If Button = 2 Then
            Paste
            Exit Sub
        End If

        If Point(x, y) = &HFFFFFE Then
            Mid$(Escolhidos, Setor + 1, 1) = "+"
            mbrush = CreateSolidBrush(6528)
            SelectObject hDC, mbrush
            ScaleMode = vbPixels
            ExtFloodFill hDC, x, y, GetPixel(hDC, x, y), FLOODFILLSURFACE
            DeleteObject mbrush
            Refresh
          Else
            Mid$(Escolhidos, Setor + 1, 1) = "-"
            mbrush = CreateSolidBrush(&HFFFFFE)
            SelectObject hDC, mbrush
            ScaleMode = vbPixels
            ExtFloodFill hDC, x, y, GetPixel(hDC, x, y), FLOODFILLSURFACE
            DeleteObject mbrush
            Refresh
        End If
        Exit Sub
    End If
   
    nMousePosition = MakeDWord(x, y)
    Select Case Button
      Case 1
        PostMessage Form1.hwnd, WM_LBUTTONDOWN, Button, nMousePosition
      Case 2
        PostMessage Form1.hwnd, WM_RBUTTONDOWN, Button, nMousePosition
    End Select

End Sub

Public Function MakeDWord(LoWord As Single, HiWord As Single) As Long

    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)

End Function

Private Sub Prepare_Mascara()

  Static Raio As Integer
  Dim i As Double
  Dim x As Single
  Dim y As Single
  Dim mbrush As Long
  
    Raio = 208
    'FillStyle = 0
    Picture1.ForeColor = &HA5B423
    Passo = 5.625 * 4
    Picture1.Circle (CentroX, CentroY), Raio
    For i = 0 To 16
        ang = ang + Passo

        ag = Radiano * ang
        xt = (Sin(ag)) * Raio
        yt = -Cos(ag) * Raio
        Picture1.Line (CentroX + xt - Raio * Sin(ag), CentroY + yt - Raio * -Cos(ag))-(CentroX + xt, CentroY + yt)
    Next i

    ang = 360 - 20
    Raio = 100
    For i = 0 To 15
    
        ang = ang + Passo
        ag = Radiano * ang
        xt = (Sin(ag)) * Raio
        yt = -Cos(ag) * Raio
        x = CentroX + xt
        y = CentroY + yt
            
        mbrush = CreateSolidBrush(i * 10000)
        SelectObject Picture1.hDC, mbrush
        Picture1.ScaleMode = vbPixels
        ExtFloodFill Picture1.hDC, x, y, GetPixel(Picture1.hDC, x, y), FLOODFILLSURFACE
        DeleteObject mbrush
        Picture1.Refresh
    Next i

End Sub

Private Sub Paste()

  Dim Achou As Integer, m As Integer, i As Integer, k As Integer
  
    If Form1.Shape1(0).FillColor = &HFF Then
        Form1.Capture 5.625
        m = 16
      Else
        Form1.Capture 7.5
        m = 12
    End If
    
    Prepare_Undo
        
    For i = 0 To 15
        If Mid$(Escolhidos, i + 1, 1) = "+" Then
        
            For k = 0 To m - 1
                Dados((Setor) * m + k) = Backup(i * m + k, (Nivel_para_Undo - 1) And Quant_UNDO)
            Next k
            Achou = True
        End If
        If Achou Then
            Setor = Setor + 1
            If Setor > 15 Then
                Setor = 0
            End If
        End If
    Next i

    If Form1.Shape1(0).FillColor = &HFF Then
        Form1.Pintar 5.625
      Else
        Form1.Pintar 7.5
    End If

    Backup_34ou44(Nivel_para_Undo) = Form1.Shape1(0).FillColor = &HFF
    For i = 0 To 256
        Backup(i, Nivel_para_Undo) = Dados(i)
    Next i

End Sub

Private Sub Undo()

  Dim i As Integer


    Nivel_para_Undo = (Nivel_para_Undo - 1) And Quant_UNDO
      
    If Backup(0, Nivel_para_Undo) = 0 Then
        Nivel_para_Undo = (Nivel_para_Undo + 1) And Quant_UNDO
        Exit Sub
    End If

    For i = 0 To 256
        Dados(i) = Backup(i, Nivel_para_Undo)
    Next i

    If Backup_34ou44(Nivel_para_Undo) Then
        Form1.Label2_DblClick (0)
      Else
        Form1.Label2_DblClick (1)
    End If

    If Form1.Shape1(0).FillColor = &HFF Then
        Form1.Pintar 5.625
      Else
        Form1.Pintar 7.5
    End If

End Sub

Public Sub Prepare_Undo()

  Dim i As Integer
    
    Backup_34ou44(Nivel_para_Undo) = Form1.Shape1(0).FillColor = &HFF
    For i = 0 To 256
        Backup(i, Nivel_para_Undo) = Dados(i)
    Next i
    
    Nivel_para_Undo = (Nivel_para_Undo + 1) And Quant_UNDO
    

End Sub
    
Private Sub Redo()

  Dim i As Integer

    Nivel_para_Undo = (Nivel_para_Undo + 1) And Quant_UNDO

    If Backup(0, Nivel_para_Undo) = 0 Then
        Nivel_para_Undo = (Nivel_para_Undo - 1) And Quant_UNDO
        Exit Sub
    End If
    
    For i = 0 To 256
        Dados(i) = Backup(i, Nivel_para_Undo)
    Next i

    If Backup_34ou44(Nivel_para_Undo) Then
        Form1.Label2_DblClick (0)
      Else
        Form1.Label2_DblClick (1)
    End If

    If Form1.Shape1(0).FillColor = &HFF Then
        Form1.Pintar 5.625
      Else
        Form1.Pintar 7.5
    End If

End Sub


