VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   6600
   ClientLeft      =   3945
   ClientTop       =   900
   ClientWidth     =   4860
   Icon            =   "split_sobre.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "split_sobre.frx":058A
   ScaleHeight     =   6600
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://nevermindrs.cjb.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1125
      TabIndex        =   0
      Top             =   6255
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   6585
      Left            =   0
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsFuncoes
Private Sub Form_Load()
Dim hSysMenu As Long
Set ontop = New clsFuncoes
ontop.MakeTopMost hWnd
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMM
Form2.Caption = versao & " - Sobre [Build " & buildver & "]"
Image1.Picture = LoadPicture(App.Path & "\noicons.ni4")
End Sub

Private Sub Image1_Click()
Me.Hide
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = vbBlue
Me.MousePointer = 0
End Sub

Private Sub Label1_Click()
Executar Label1.Caption, "", ""

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = vbCyan
Me.MousePointer = 99

End Sub
