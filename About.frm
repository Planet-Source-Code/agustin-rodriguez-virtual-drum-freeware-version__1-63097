VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu About 
         Caption         =   "About"
         Begin VB.Menu nome 
            Caption         =   "Agustin Rodriguez"
            Index           =   0
            Begin VB.Menu opt 
               Caption         =   "E-Mail"
               Index           =   1
            End
            Begin VB.Menu opt 
               Caption         =   "Home Page"
               Index           =   2
            End
         End
      End
      Begin VB.Menu Help 
         Caption         =   "Help"
         Begin VB.Menu hlp 
            Caption         =   "Right Button on the circle               = Mark"
            Index           =   0
         End
         Begin VB.Menu hlp 
            Caption         =   "Left Button on the circle                 = Unmark"
            Index           =   1
         End
         Begin VB.Menu hlp 
            Caption         =   "SHIFT + Right Button on Sector     = Select Sector"
            Index           =   2
         End
         Begin VB.Menu hlp 
            Caption         =   "SHIFT + Left Button on Sector       = Paste Selected Sectors"
            Index           =   3
         End
         Begin VB.Menu hlp 
            Caption         =   "ALT + Right Button                         = Undo"
            Index           =   4
         End
         Begin VB.Menu hlp 
            Caption         =   "ALT + Left Button                           = Redo"
            Index           =   5
         End
      End
      Begin VB.Menu Register 
         Caption         =   "Register"
         Begin VB.Menu reg 
            Caption         =   "       *        This is FREE        *"
            Index           =   0
         End
         Begin VB.Menu reg 
            Caption         =   "If you like and it to be useful for you"
            Index           =   1
         End
         Begin VB.Menu reg 
            Caption         =   "then  make a donation for the Autor"
            Index           =   2
         End
         Begin VB.Menu reg 
            Caption         =   "or for a Charity Institution."
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Msg As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const conSwNormal As Long = 1

Private Sub opt_Click(Index As Integer)

    On Error GoTo Erro
    Select Case Index
      Case 1
        ShellExecute hwnd, "open", "mailto:virtual_guitar_1@hotmail.com", vbNullString, vbNullString, conSwNormal
      Case 2
        ShellExecute hwnd, "open", "http://geocities.com/virtual_quality/", vbNullString, vbNullString, conSwNormal
    End Select
      
Erro:
    Resume Next

End Sub


