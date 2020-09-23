VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMail 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PopCheck"
   ClientHeight    =   2115
   ClientLeft      =   2625
   ClientTop       =   2130
   ClientWidth     =   6750
   Icon            =   "vemail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6750
   Begin NoIcons43.Button3D Command5 
      Height          =   330
      Left            =   5490
      TabIndex        =   5
      Top             =   1710
      Width           =   1140
      _ExtentX        =   2011
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
   Begin NoIcons43.Button3D Command2 
      Height          =   330
      Left            =   4275
      TabIndex        =   4
      Top             =   1710
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      BackColor       =   0
      Caption         =   "Desconectar"
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
      Left            =   3060
      TabIndex        =   3
      Top             =   1710
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      BackColor       =   0
      Caption         =   "Conectar"
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
      Left            =   1845
      TabIndex        =   2
      Top             =   1710
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      BackColor       =   0
      Caption         =   "Configurar"
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   45
      Top             =   1260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lmails 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000FF00&
      Height          =   1035
      Left            =   90
      TabIndex        =   1
      Top             =   225
      Width           =   6540
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5265
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2250
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3555
      TabIndex        =   0
      Top             =   1305
      Width           =   3075
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim response, acao, vok
Public username, password, nmsg, smtp

Private Sub Command1_Click()
lmails.Clear
Data1.RecordSource = "select distinct * from pop"
Data1.Refresh
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
    On Error Resume Next
    checkar Data1.Recordset!smtp
    Data1.Recordset.MoveNext
Loop
Call Form_Load
Command1.Enabled = True

stat "Fim"
End Sub

Private Sub Command2_Click()
If Winsock1.State <> sckClosed Then
    Winsock1.Close
    stat "Desconectando..."
    Command2.Enabled = False
    Command1.Enabled = True
End If

End Sub

Private Sub Command3_Click()
Me.Enabled = False
Form5.Show

End Sub

Private Sub Command4_Click()
If Winsock1.State <> sckClosed Then Winsock1.Close
Call Form_Load
stat "Operação cancelada"
Command1.Enabled = True

End Sub

Private Sub Command5_Click()
If Winsock1.State <> sckClosed Then Winsock1.Close
Unload Me
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMM

Data1.DatabaseName = App.Path & "\noicons.mdb"
Data1.RecordSource = "select * from pop"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
smtp = Data1.Recordset!smtp
username = Data1.Recordset!username
password = Data1.Recordset!password
Command2.Enabled = False
stat ""

End Sub
Private Sub wait(cod As String)
Do Until Trim(Mid(response, 1, 3)) = Trim(cod)
    If acao = "pass" And Trim(Mid(response, 1, 1)) = "-" Then
        acao = "erro"
        Exit Sub
    End If
    If Trim(Mid(response, 1, 4)) = "-ERR" Then Exit Sub
    DoEvents

Loop
End Sub
Private Sub stat(log As String)
Label1.Caption = log
End Sub
Private Sub checkar(sserver As String)
vok = ""
smtp = sserver
username = Data1.Recordset!username
password = Data1.Recordset!password

If smtp = "" Or username = "" Or password = "" Then
    x = MsgBox("As configurações de e-mail estão incorretas." & Chr(10) & "Deseja configurar seu e-mail agora?", vbExclamation + vbYesNo, "PopCheck - Configuração")
    If x = vbYes Then Command3_Click
    Command2_Click
    Exit Sub
End If

If Not online Then
    MsgBox "Você não está conectado"
    Exit Sub
End If
Command1.Enabled = False
Command2.Enabled = True
If Winsock1.State <> sckClosed Then Winsock1.Close
stat "Conectando ao servidor " & smtp
Winsock1.Connect smtp, 110
Do While vok = ""
    DoEvents
Loop
End Sub


Private Sub Winsock1_Connect()
response = ""
wait "+OK"
response = ""
stat "Efetuando login..."
Winsock1.SendData "user " & username & vbCrLf
wait "+OK"
response = ""
acao = "pass"
Winsock1.SendData "pass " & password & vbCrLf
wait "+OK"
If acao = "erro" Then
    Command2_Click
    stat "Login incorreto - " & sserver
    acao = ""
    lmails.AddItem "[" & username & " - " & smtp & "] - Login inválido [ERRO]"
    acao = ""
    vok = "ok"
    Exit Sub
End If

For i = 1 To 1000000
Next
response = ""
acao = "stat"
stat "Verificando e-mail..."
Winsock1.SendData "stat " & vbCrLf
wait "+OK"
If nmsg = 0 Then lmails.AddItem "[" & username & " - " & smtp & "] - Nenhuma mensagem nova"
If nmsg = 1 Then lmails.AddItem "[" & username & " - " & smtp & "] - 1 mensagem nova"
If nmsg > 1 Then lmails.AddItem "[" & username & " - " & smtp & "] - " & nmsg & " mensagens novas"
vok = "ok"
Winsock1.Close

End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData response, vbString
If acao = "stat" Then
nmsg = Trim(Mid(response, 5, 2))
End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Number = 11053 Then
    stat "Conexão com o servidor falhou."
End If
Command2_Click
stat "Conexão com o servidor falhou."
    lmails.AddItem "[" & username & " - " & smtp & "] - Não foi possível conectar ao servidor."
    vok = "ok"
End Sub
