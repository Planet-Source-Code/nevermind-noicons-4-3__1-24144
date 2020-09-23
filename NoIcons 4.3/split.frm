VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "NoIcons 4.3"
   ClientHeight    =   405
   ClientLeft      =   2790
   ClientTop       =   3780
   ClientWidth     =   2610
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "split.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   405
   ScaleWidth      =   2610
   Visible         =   0   'False
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   -360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   90
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   2610
      Top             =   90
   End
   Begin MSComctlLib.ImageList img 
      Left            =   1215
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "split.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "split.frx":0B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "split.frx":0F7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Menu mn_menu_principal 
      Caption         =   "Menu"
      Begin VB.Menu mn_item 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu dsad 
         Caption         =   "-"
      End
      Begin VB.Menu noicons 
         Caption         =   "[ NoIcons ]"
         Begin VB.Menu pop 
            Caption         =   "[ PopCheck ]"
         End
         Begin VB.Menu agenda 
            Caption         =   "[ Agenda ]"
         End
         Begin VB.Menu qpad 
            Caption         =   "[ QuickPad ]"
         End
         Begin VB.Menu infohd2 
            Caption         =   "[ InfoHd ]"
         End
         Begin VB.Menu f 
            Caption         =   "-"
         End
         Begin VB.Menu config 
            Caption         =   "[ Configuração ]"
         End
         Begin VB.Menu alticon 
            Caption         =   "[ Alterar Ícone ]"
         End
         Begin VB.Menu sobre 
            Caption         =   "[ Sobre... ]"
         End
         Begin VB.Menu sair 
            Caption         =   "[ Sair ]"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I would like to tanks all the people that submit their
'codes to planet-source-code.com!
Dim str_win_command(1 To 100) As String
Public WithEvents SysIcon As clsFuncoes
Attribute SysIcon.VB_VarHelpID = -1
Dim i As Integer

Private Sub agenda_Click()
Form3.Show

End Sub

Private Sub alticon_Click()
frmIcon.Visible = True

End Sub

Private Sub config_Click()
titlebar.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
       Static lngMsg As Long
       Static blnFlag As Boolean
       Dim result As Long
       
       lngMsg = x / Screen.TwipsPerPixelX


       If blnFlag = False Then
           blnFlag = True


           Select Case lngMsg
               'doubleclick
               Case WM_LBUTTONDBLCLICK
               'Form2.Show
               'right-click
               Case WM_RBUTTONUP
               result = SetForegroundWindow(Me.hWnd)
               Me.PopupMenu mn_menu_principal
'               Me.PopupMenu mnuSystemTray
           End Select


       blnFlag = False
   End If


   End Sub
Public Function changeicon()
    SysIcon.IconHandle = Me.Icon
End Function

Private Sub Form_Load()
versao = "NoIcons 4.3"
buildver = "04/06/01"
'CheckTask "NoIcons [4.3]"
If App.PrevInstance Then
    MsgBox "O programa já está aberto."
    End
End If

Me.Visible = False
icone = 1
Form1.Caption = versao
If Not FunMontaMenu Then
    he = "he"
End If
Data1.RecordSource = "select * from agenda where data like '*" & Date & "*'"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
    icone = 2
    For i = 0 To 5
        Beep
    Next
End If

Data2.DatabaseName = App.Path & "\noicons.mdb"
Data2.RecordSource = "select * from icon"
Data2.Refresh
If Data2.Recordset!current = "Default" Then
    Me.Icon = Form1.img.ListImages(1).Picture
Else
    Me.Icon = LoadPicture(Data2.Recordset!Path)
End If




Set SysIcon = New clsFuncoes
SysIcon.Initialize hWnd, Form1.Icon, versao & " feito por «× N e v e R m i n D ×»"
SysIcon.ShowIcon
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SysIcon.HideIcon
End Sub

Private Sub infohd2_Click()
InfoHd.Show
End Sub

Private Sub pop_Click()
frmMail.Show

End Sub

Private Sub qpad_Click()
quickpad.Show
End Sub

Private Sub sair_Click()
SysIcon.HideIcon
End
End Sub
Private Sub sobre_Click()
Form2.Show
End Sub

Public Function FunMontaMenu() As Boolean
If Dir(App.Path & "\noicons.mdb") = "" Then
    MsgBox "Não foi possível localizar banco de dados." & Chr(10) & "Verifique se ele se encontra no mesmo diretório do executável." & Chr(10) & "Leia o arquivo Readme.txt para maiores informações", vbOKOnly + vbCritical, versao
    End
End If
Data1.DatabaseName = App.Path & "\noicons.mdb"
Data1.RecordSource = "select * from noicons"
Data1.Refresh
If Data1.Recordset.EOF Then
    MsgBox "Não há nada para adicionar.", vbExclamation + vbOKOnly, "NoIcons"
    titlebar.Show
    Exit Function
       
End If
For i = 1 To 99
On Error Resume Next
    Unload mn_item(i)
    mn_item(i).Visible = False
Next
i = 1
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
        
    If remonta = "sim" Then
        If mn_item(i).Caption <> "" Then
            Unload mn_item(i)
        End If
    End If
    
    Load mn_item(i)
    mn_item(i).Caption = Data1.Recordset!nome
    str_win_command(i) = Data1.Recordset!Path & "||" & Data1.Recordset!opcoes
    mn_item(i).Visible = True
    i = i + 1
        Data1.Recordset.MoveNext

Loop
mn_item(0).Visible = False
mn_menu_principal.Visible = True
FunMontaMenu = True
remonta = ""
End Function

Private Sub mn_item_Click(Index As Integer)
temp = str_win_command(Index)
Path = Mid(str_win_command(Index), 1, InStr(1, str_win_command(Index), "||") - 1)
param = Mid(temp, Len(Path) + 3, Len(temp))
i = 1
For i = i To Len(Path)
If Mid(Path, i, 1) = "\" Then
    tpos = i
End If
Next
If InStr(1, Path, "://") = 0 Then ChDir Mid(Path, 1, tpos)
If Trim(param) = "" Then
x = Executar(Path, "", "")
Else
x = Executar(Path, param, "")
End If
End Sub

Private Sub Timer1_Timer()
Data1.RecordSource = "select * from agenda where data like '*" & Date & "*'"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
    If icone = 1 Then
        icone = 2
        SysIcon.IconHandle = img.ListImages(icone).Picture
        For i = 0 To 5
        Beep
        Next
    End If
Else
    If icone = 2 Then
        icone = 1
        SysIcon.IconHandle = img.ListImages(icone).Picture
    End If
End If

End Sub
