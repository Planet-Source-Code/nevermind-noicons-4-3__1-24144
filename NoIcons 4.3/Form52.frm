VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PopCheck - Configuração"
   ClientHeight    =   2340
   ClientLeft      =   3540
   ClientTop       =   2280
   ClientWidth     =   4995
   Icon            =   "Form52.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4995
   Begin NoIcons43.Button3D Command5 
      Height          =   330
      Left            =   3960
      TabIndex        =   12
      Top             =   1890
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      BackColor       =   0
      Caption         =   "Fechar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin NoIcons43.Button3D Command4 
      Height          =   330
      Left            =   2925
      TabIndex        =   11
      Top             =   1890
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      BackColor       =   0
      Caption         =   "Remover"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin NoIcons43.Button3D Command3 
      Height          =   330
      Left            =   1890
      TabIndex        =   10
      Top             =   1890
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      BackColor       =   0
      Caption         =   "Nova Conta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin NoIcons43.Button3D Command2 
      Height          =   330
      Left            =   990
      TabIndex        =   9
      Top             =   1890
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   582
      BackColor       =   0
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin NoIcons43.Button3D Command1 
      Height          =   330
      Left            =   90
      TabIndex        =   8
      Top             =   1890
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   582
      BackColor       =   0
      Caption         =   "Atualizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   2940
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "§"
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1485
      TabIndex        =   4
      Top             =   1080
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   900
      TabIndex        =   3
      Top             =   720
      Width           =   2265
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2970
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Conta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   270
      TabIndex        =   7
      Top             =   405
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   315
      TabIndex        =   2
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   315
      TabIndex        =   1
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Pop:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   270
      TabIndex        =   0
      Top             =   720
      Width           =   555
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()

Data1.RecordSource = "select * from pop where smtp = '" & Mid(Combo1.text, InStr(1, Combo1.text, "-") + 2, Len(Combo1.text)) & "' and username = '" & Mid(Combo1.text, 1, InStr(1, Combo1.text, "-") - 2) & "'"
Data1.Refresh
If Data1.Recordset.EOF Then
    MsgBox "Não foi possível localizar esta conta no banco de dados.", vbOKOnly + vbInformation, "PopCheck"
    Exit Sub
End If
Text1.text = Data1.Recordset!smtp
Text2.text = Data1.Recordset!username
Text3.text = Data1.Recordset!password

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Command1_Click()
If Command1.Caption = "Adicionar" Then
    Data1.RecordSource = "select * from pop"
    Data1.Refresh
    Data1.Recordset.AddNew
    Data1.Recordset!smtp = Text1
    Data1.Recordset!username = Text2
    Data1.Recordset!password = Text3
    Data1.Recordset.Update
    Command3.Enabled = True
    Command2.Enabled = False
    Command1.Caption = "Atualizar"
    Command4.Enabled = True
    Call Form_Load
End If

If Command1.Caption = "Atualizar" Then
    smtp = Text1.text
    username = Text2.text
    password = Text3.text
    Data1.RecordSource = "select * from pop where smtp = '" & Mid(Combo1.text, InStr(1, Combo1.text, "-") + 2, Len(Combo1.text)) & "' and username = '" & Mid(Combo1.text, 1, InStr(1, Combo1.text, "-") - 2) & "'"
    Data1.Refresh
    Data1.Recordset.Edit
    Data1.Recordset!smtp = Text1.text
    Data1.Recordset!username = Text2.text
    Data1.Recordset!password = Text3.text
    Data1.Recordset.Update
    Call Form_Load
End If
End Sub

Private Sub Command2_Click()
If Command1.Caption = "Adicionar" Then
    Combo1.ListIndex = 0
    Command3.Enabled = True
    Command2.Enabled = False
    Command4.Enabled = True
    Command1.Caption = "Atualizar"
    Exit Sub
End If



End Sub

Private Sub Command3_Click()
Command1.Caption = "Adicionar"
Command3.Enabled = False
Command4.Enabled = False
Command2.Enabled = True
Text1 = ""
Text2 = ""
Text3 = ""

End Sub

Private Sub Command4_Click()
On Error Resume Next
x = MsgBox("Você tem certeza que deseja excluir a conta " & Combo1.text & "?", vbYesNo + vbExclamation, "PopCheck - Excluir")
If x = vbNo Then Exit Sub
Data1.RecordSource = "select * from pop where smtp = '" & Mid(Combo1.text, InStr(1, Combo1.text, "-") + 2, Len(Combo1.text)) & "' and username = '" & Mid(Combo1.text, 1, InStr(1, Combo1.text, "-") - 2) & "'"
Data1.Refresh
Data1.Recordset.Delete
remonta
End Sub

Private Sub Command5_Click()
frmMail.Enabled = True
Me.Visible = False
frmMail.SetFocus
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMM
Combo1.Clear
remonta

End Sub
Private Sub remonta()
Combo1.Clear

Data1.DatabaseName = App.Path & "\noicons.mdb"
Data1.RecordSource = "select * from pop"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
Do While Not Data1.Recordset.EOF
    Combo1.AddItem Data1.Recordset!username & " - " & Data1.Recordset!smtp
    Data1.Recordset.MoveNext
Loop
Combo1.ListIndex = 0
End Sub
Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub


Private Sub Text2_GotFocus()
Text2.SelStart = 0
Text2.SelLength = Len(Text2)
End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3)
End Sub
