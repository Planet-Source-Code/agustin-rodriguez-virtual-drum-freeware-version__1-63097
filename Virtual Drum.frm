VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000001&
   Icon            =   "Virtual Drum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Virtual Drum.frx":164A
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   600
      Top             =   6720
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3120
      MouseIcon       =   "Virtual Drum.frx":AEA1C
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   6960
      MouseIcon       =   "Virtual Drum.frx":AED26
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "3/4"
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   1
      Left            =   4200
      MouseIcon       =   "Virtual Drum.frx":AF030
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6690
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4/4"
      ForeColor       =   &H00000001&
      Height          =   255
      Index           =   0
      Left            =   2520
      MouseIcon       =   "Virtual Drum.frx":AF33A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6690
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      X1              =   240
      X2              =   240
      Y1              =   244
      Y2              =   40
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Play"
      ForeColor       =   &H00000001&
      Height          =   195
      Left            =   3420
      MouseIcon       =   "Virtual Drum.frx":AF644
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   7020
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents XC As Xcom
Attribute XC.VB_VarHelpID = -1

Option Explicit

Private Sub Form_Load()

  'RUN ONLY ON WINDOWS XP
                   
  'Shut Down using the X from the Form2 , NOT VB IDE

  Dim NormalWindowStyle As Long
  Dim ret As Long, xx As Integer, n As String, i As Integer
  Dim col As Long
  
    Set XC = New Xcom
    XC.Start "Virtual Metronome XCOM"
    XC.Connect ("VIRTUAL GUITAR x VIRTUAL METRONOME")
    XC.Send "METRONOME ATIVADO"
  
    '================================
    'USE OTHER MIDI DEVICE HERE
    Dev_OUT = 0
    '================================
  
    Form3.Show , Me
    
    CentroX = 240
    CentroY = 240
    Raio = 204
    
    Divisão = 2
    
    Timer1.Interval = 60000 / ((2 ^ Divisão) * 120)

    Open App.Path + "\drumpth.ini" For Input As 1
leia_outro:
    Do While Not EOF(1) = -1
        Input #1, n
        If n = "" Then
            GoTo leia_outro
        End If
    
        If left$(n, 1) = "[" Then
            xx = Val(Mid$(n, 2, 3))
            Bank_util(xx) = True
            For i = 0 To 127
                Input #1, n
                Drum_Name(i, xx) = Mid$(n, 5)
            Next i
        End If
    Loop

va_em_frente:
    Close
    For i = 0 To 7
        Form2.Inst_name(i) = Drum_Name(Form2.Label3(i), 0)
    Next i

    'SetOnTop Me, True
    
    MidiOpen
    
    Cor_atual = 255
 
    Radiano = 3.14156 / 180
 
    NormalWindowStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 50, LWA_ALPHA

    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    col = RGB(0, 0, 0)
    SetLayeredWindowAttributes Me.hwnd, col, 50, LWA_COLORKEY

    Obj(0).tamanho = 9
    Obj(0).cor = Form2.Shape1(0).FillColor
    Obj(0).Raio = 200

    Obj(1).tamanho = 8
    Obj(1).cor = Form2.Shape1(1).FillColor
    Obj(1).Raio = 180

    Obj(2).tamanho = 7
    Obj(2).cor = Form2.Shape1(2).FillColor
    Obj(2).Raio = 162

    Obj(3).tamanho = 6
    Obj(3).cor = Form2.Shape1(3).FillColor
    Obj(3).Raio = 147

    Prepare 5.625

End Sub

Private Sub Mova_Ponteiro(Passo)

  Dim i As Integer

    ang = ang + Passo
    If ang = 360 Then
        ang = 0
    End If

    ag = Radiano * ang
    xt = (Sin(ag)) * Raio
    yt = -Cos(ag) * Raio
    Line1.X1 = CentroX + xt - Raio * Sin(ag)
    Line1.Y1 = CentroY + yt - Raio * -Cos(ag)
    Line1.X2 = CentroX + xt
    Line1.Y2 = CentroY + yt

    For i = 0 To 3
        xt = (Sin(ag)) * Obj(i).Raio
        yt = -Cos(ag) * Obj(i).Raio
        
        Select Case Point(CentroX + xt, CentroY + yt)
          Case 255
            ShortMessage &H99, Form2.Label3(0), 127
          Case 65280
            ShortMessage &H99, Form2.Label3(1), 127
          Case 65535
            ShortMessage &H99, Form2.Label3(2), 127
          Case 16776960
            ShortMessage &H99, Form2.Label3(3), 127
          Case 16761024
            ShortMessage &H99, Form2.Label3(4), 127
          Case 33023
            ShortMessage &H99, Form2.Label3(5), 127
          Case 16711935
            ShortMessage &H99, Form2.Label3(6), 127
          Case 16711680
            ShortMessage &H99, Form2.Label3(7), 127
        End Select
            
    Next i

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim ret As Long, mbrush As Long, i As Integer

    ret = Point(x, y)
    
    If Button = 1 Then
        For i = 0 To 7
            If Point(x, y) = Form2.Shape1(i).FillColor Then
                GoTo ok
            End If
                
        Next i
        If Point(x, y) <> vbWhite Then
            Exit Sub
        End If
            
ok:
        mbrush = CreateSolidBrush(Cor_atual)
        SelectObject hDC, mbrush
        ScaleMode = vbPixels
        ExtFloodFill hDC, x, y, GetPixel(hDC, x, y), FLOODFILLSURFACE
        DeleteObject mbrush
      Else
        For i = 0 To 7
            If Point(x, y) = Form2.Shape1(i).FillColor Then
                GoTo ok1
            End If
                
        Next i
        If Point(x, y) <> vbWhite Then
            Exit Sub
        End If
            
ok1:
        mbrush = CreateSolidBrush(16777215)
        SelectObject hDC, mbrush
        ScaleMode = vbPixels
        ExtFloodFill hDC, x, y, GetPixel(hDC, x, y), FLOODFILLSURFACE
        DeleteObject mbrush

    End If

End Sub

Private Sub Form_Resize()

    Form3.Show , Me
    Form2.Move left + Width + 10000

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If XC.Connected Then
        XC.Send "METRONOME DESATIVADO"
    End If

End Sub

Private Sub Image1_Click()

  Static vez As Integer, i As Integer

    vez = vez Xor -1
    For i = 0 To 100
        DoEvents
        If vez Then
            If Form2.left = 6925 Then
                Exit Sub
            End If
                
            Form2.Move Form2.left - 50
          Else
            If Form2.left = 11975 Then
                Exit Sub
            End If
                
            Form2.Move Form2.left + 50
        End If
    Next i

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Shape1(0).FillColor = &HFF Then
        ang = 360 - 5.625
      Else
        ang = 360 - 7.5
    End If
           
    Timer1.Enabled = Timer1.Enabled Xor -1

    If Timer1.Enabled Then
        Label3 = "Stop"
      Else
        Label3 = "Play"
    End If

End Sub

Public Sub Label2_DblClick(Index As Integer)

    Shape1(Index).FillColor = &HFF
    Shape1(Index Xor 1).FillColor = &H808080
    Cls
    If Index Then
        Prepare 7.5
      Else
        Prepare 5.625
    End If

End Sub

Private Sub Timer1_Timer()

    If Shape1(0).FillColor = &HFF Then
        Mova_Ponteiro 5.625
      Else
        Mova_Ponteiro 7.5
    End If

End Sub

Private Sub Put_Circle(x, y, size, cor As Long)

    FillColor = cor
    Circle (x, y), size

End Sub

Private Sub Prepare(Passo)

  Dim m As Integer, i As Integer, Quant As Integer

    ang = 0
    
    If Passo = 7.5 Then
        Quant = 47
      Else
        Quant = 63
    End If
        
    For m = 0 To Quant
        ang = ang + Passo
    
        If ang = 360 Then
            ang = 0
        End If
            
        ag = Radiano * ang
        For i = 0 To 3
            xt = (Sin(ag)) * Obj(i).Raio
            yt = -Cos(ag) * Obj(i).Raio
            Put_Circle CentroX + xt, CentroY + yt, Obj(i).tamanho, &HFFFFFF
        Next i
    Next m

End Sub

Public Sub Capture(Passo)
  
  Dim qt As Integer, m As Integer, i As Integer, Quant As Integer

    ang = 0
    If Passo = 7.5 Then
        Quant = 47
      Else
        Quant = 63
    End If
    
    For m = 0 To Quant
                    
        ag = Radiano * ang
        For i = 0 To 3
            xt = (Sin(ag)) * Obj(i).Raio
            yt = -Cos(ag) * Obj(i).Raio
            Dados(qt) = Point(CentroX + xt, CentroY + yt)
            qt = qt + 1
        Next i
        ang = ang + Passo
    Next m

End Sub

Public Sub Pintar(Passo)

  Dim m As Integer, i As Integer, qt As Integer, Quant As Integer
  
    ang = 0
    
    If Passo = 7.5 Then
        Quant = 47
      Else
        Quant = 63
    End If
        
    For m = 0 To Quant
        ag = Radiano * ang
        xt = (Sin(ag)) * Raio
        yt = -Cos(ag) * Raio
        For i = 0 To 3
            xt = (Sin(ag)) * Obj(i).Raio
            yt = -Cos(ag) * Obj(i).Raio
            Cor_atual = Dados(qt)
            qt = qt + 1
            Form_MouseDown 1, 0, CentroX + xt, CentroY + yt
        Next i
        ang = ang + Passo
    Next m

End Sub

Private Sub XC_DataReceived(data As String)

    If data = "[START]" Then
        If Shape1(0).FillColor = &HFF Then
            ang = 360 - 5.625
          Else
            ang = 360 - 7.5
        End If
        Timer1.Enabled = True
        Label3 = "Stop"
    End If

    If left(data, 4) = "BPM=" Then
        Form2.Label1 = Val(Mid$(data, 5))
        Form1.Timer1.Interval = 60000 / ((2 ^ Divisão) * Form2.Label1)
    End If
    
    If left(data, 6) = "TICKS=" Then
        '
    End If
    
    If data = "[STOP]" Then
        Timer1.Enabled = False
        Label3 = "Play"
    End If

    If data = "VIRTUAL GUITAR ATIVADO" Then
        XC.Connect ("VIRTUAL GUITAR x VIRTUAL METRONOME")
        XC.Send "METRONOME ATIVADO"
    End If
    
    If data = "VIRTUAL GUITAR DESLIGADO" Then
        Xcom.Quit
    End If

End Sub


