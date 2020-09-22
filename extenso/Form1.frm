VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   8295
   End
   Begin VB.Label Label2 
      Caption         =   "Valor em numerário"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Valor por extenso"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Activate()
Text2.SetFocus
End Sub

Private Sub Text2_Change()
On Error Resume Next
Dim sRet As String
Dim dValor As Double
dValor = Text2
sRet = Extenso(dValor, "Euros", "Euro")
Text1.Text = sRet

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    KeyAscii = 44
    End If
End Sub
