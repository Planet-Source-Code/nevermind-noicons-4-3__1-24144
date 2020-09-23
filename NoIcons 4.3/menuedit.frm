VERSION 5.00
Begin VB.Form titlebar 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3810
   ClientLeft      =   960
   ClientTop       =   780
   ClientWidth     =   9375
   Icon            =   "menuedit.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   9375
   Begin NoIcons43.Button3D Button3D2 
      Height          =   240
      Left            =   315
      TabIndex        =   11
      Top             =   1755
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackColor       =   0
      Caption         =   "t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin NoIcons43.Button3D Button3D1 
      Height          =   240
      Left            =   315
      TabIndex        =   10
      Top             =   2070
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackColor       =   0
      Caption         =   "u"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65280
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   5895
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox list 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000FF00&
      Height          =   3180
      Left            =   675
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   495
      Width           =   3120
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000012&
      Caption         =   "Carregar NoIcons quando o windows for iniciado."
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   675
      TabIndex        =   5
      ToolTipText     =   "Adicionar NoIcons ao registro para iniciar junto com o windows."
      Top             =   225
      Width           =   3840
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -1890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5805
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame frm_add 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   2205
      Left            =   3870
      TabIndex        =   0
      Top             =   405
      Width           =   5400
      Begin NoIcons43.Button3D Command8 
         Height          =   375
         Left            =   765
         TabIndex        =   17
         Top             =   1665
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
         BackColor       =   0
         Caption         =   "Novo"
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
      Begin NoIcons43.Button3D Command7 
         Height          =   375
         Left            =   765
         TabIndex        =   18
         Top             =   1665
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
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
      Begin NoIcons43.Button3D Command5 
         Height          =   375
         Left            =   2025
         TabIndex        =   20
         Top             =   1665
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
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
      Begin NoIcons43.Button3D Command1 
         Height          =   375
         Left            =   3285
         TabIndex        =   19
         Top             =   1665
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
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
         Height          =   285
         Left            =   4635
         TabIndex        =   16
         Top             =   810
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         BackColor       =   0
         Caption         =   "ClipBrd"
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
         Height          =   285
         Left            =   4275
         TabIndex        =   15
         Top             =   810
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         BackColor       =   0
         Caption         =   "..."
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
      Begin NoIcons43.Button3D Command11 
         Height          =   330
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         BackColor       =   0
         Caption         =   "Separador"
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
      Begin VB.TextBox Text3 
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1170
         TabIndex        =   7
         Top             =   1170
         Width           =   1860
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   780
         Width           =   3390
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   990
         TabIndex        =   3
         Top             =   360
         Width           =   3240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opções: "
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
         Left            =   225
         TabIndex        =   6
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
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
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   765
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   1095
      Left            =   7785
      TabIndex        =   9
      Top             =   2610
      Width           =   1500
      Begin NoIcons43.Button3D Button3D4 
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   630
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
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
      Begin NoIcons43.Button3D Button3D3 
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   180
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
         BackColor       =   0
         Caption         =   "Aplicar"
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
   End
End
Attribute VB_Name = "titlebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str_win_command(1 To 100) As String
Public salvar
Dim menus(0 To 100) As String
Dim menus2(0 To 100) As String
Dim mntemp(0 To 2) As String
Dim Total

Private Sub Check2_Click()
If Check1 = 1 Then
    Call savestring(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "NoIcons", App.Path & "\" & App.EXEName & ".exe")
Else
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "NoIcons")
End If

End Sub

Private Sub Button3D1_Click()
If list.ListIndex = list.ListCount - 1 Then Exit Sub
ponto = list.ListIndex + 1
temp = menus(list.ListIndex)
temp2 = menus(list.ListIndex + 1)
menus(list.ListIndex) = temp2
menus(list.ListIndex + 1) = temp
recarregar
list.ListIndex = ponto
salvou = "nao"
End Sub

Private Sub Button3D2_Click()
If list.ListIndex = 0 Then Exit Sub
ponto = list.ListIndex - 1
temp = menus(list.ListIndex)
temp2 = menus(list.ListIndex - 1)
menus(list.ListIndex) = temp2
menus(list.ListIndex - 1) = temp
recarregar
list.ListIndex = ponto
salvou = "nao"
End Sub

Private Sub Button3D3_Click()
Data1.RecordSource = "select * from noicons"
Data1.Refresh
Do While Not Data1.Recordset.EOF
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
Loop

For i = 0 To list.ListCount - 1
    If menus(i) <> "" Then
        separar menus(i)
        Data1.Recordset.AddNew
        Data1.Recordset!nome = mntemp(0)
        Data1.Recordset!Path = mntemp(1)
        Data1.Recordset!opcoes = mntemp(2)
        Data1.Recordset.Update
    End If
Next
list.Visible = True
remonta = "sim"
recarregar
Form1.FunMontaMenu
salvar = "sim"
End Sub

Private Sub Button3D4_Click()
If salvar <> "sim" Then
    x = MsgBox("Deseja salvar as alterações?", vbYesNoCancel + vbQuestion, versao)
    If x = vbYes Then Button3D3_Click
    If x = vbCancel Then Exit Sub
End If
Me.Hide
remonta = "sim"
Form1.FunMontaMenu
End Sub

Private Sub Check1_Click()
If Check1 = 1 Then
    Call savestring(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "NoIcons", App.Path & "\" & App.EXEName & ".exe")
Else
    Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "NoIcons")
End If
End Sub

Private Sub Command1_Click()
If list.text = "-" Then
x = MsgBox("Deseja realmente remover o 'Separador'?", vbQuestion + vbYesNo, versao)
Else
x = MsgBox("Deseja realmente remover o '" & list.text & "'?", vbQuestion + vbYesNo, versao)
End If
If x = vbNo Then Exit Sub
menus(list.ListIndex) = ""
list.RemoveItem list.ListIndex
'Data1.Refresh
recarregar2
recarregar
list.ListIndex = 0
End Sub



Private Sub Command11_Click()
Text1 = "-"
Text2 = "[ SEPARADOR ]"

End Sub



Private Sub Command3_Click()
x = Trim(Clipboard.GetText)
If x = "" Then
    MsgBox "Não tem nada na Área de Transferência para ser copiada", vbOKOnly, "NoIcons - Clipboard"
    Exit Sub
End If
Text2.text = CStr(x)
End Sub

Private Sub Command4_Click()
Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    On Error GoTo e_Trap
    
    FileDialog.sFilter = "Executáveis (*.exe;*.bat;*.com)" & Chr$(0) & "*.exe;*.bat;*.com" & Chr$(0) & "Arquivos Mp3 (*.mp3;*.m3u)" & Chr$(0) & "*.mp3;*.m3u" & Chr$(0) & "Todos os Arquivos (*.*)" & Chr$(0) & "*.*" & Chr$(0)
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = "Show Open"
    FileDialog.sInitDir = App.Path & "\"
    sOpen = ShowOpen(Me.hWnd)
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        Text1 = sOpen.sFiles(1)
        Text2 = sOpen.sLastDirectory & sOpen.sFiles(1)
    End If
    Exit Sub
e_Trap:
    Exit Sub
    Resume
End Sub

Private Sub Command5_Click()
If Trim(Text1.text) = "" Or Trim(Text2.text) = "" Then
    MsgBox "Não deixe nenhum dos dois primeiros campos em branco", vbOKOnly, "NoIcons - Adicionar"
    Exit Sub
End If

If Command5.Caption = "Atualizar" Then
    menus(list.ListIndex) = Text1 & "||" & Text2 & "||" & Text3
    recarregar
End If
If Command5.Caption = "Adicionar" Then
    menus(99) = Text1 & "||" & Text2 & "||" & Text3
    recarregar2
    recarregar
    list.ListIndex = list.ListCount - 1
    Command5.Caption = "Atualizar"
    Command8.Visible = True
    Command1.Visible = True
    Command7.Visible = False
    Command11.Visible = False
End If
salvou = "nao"
End Sub
Private Sub recarregar2()

'limpa a matrix temporaria
    For i = 0 To 99
        menus2(i) = ""
    Next
'passa tudo da matrix principal para a temporaria
    For i = 0 To 99
        If menus(i) <> "" Then menus2(i) = menus(i)
    Next
'limpa a matrix principal
    For i = 0 To 99
        menus(i) = ""
    Next
'passa da temporaria para a principal
x = 0
    For i = 0 To 99
        If menus2(i) <> "" Then
            menus(x) = menus2(i)
            x = x + 1
        End If
    Next
End Sub


Private Sub Command7_Click()
Command5.Caption = "Atualizar"
Command8.Visible = True
Command1.Visible = True
Command7.Visible = False
Command11.Visible = False
list.ListIndex = 0
list_Click

End Sub

Private Sub Command8_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Command5.Caption = "Adicionar"
Command8.Visible = False
Command1.Visible = False
Command7.Visible = True
Command11.Visible = True
End Sub

Private Sub Form_Load()
titlebar.Caption = versao & " - Configuração"
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMM
x = getdword(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "NoIcons")
If x = "" Then
    Check1 = 0
Else
    Check1 = 1
End If


Data1.DatabaseName = App.Path & "\noicons.mdb"
Data1.RecordSource = "select * from noicons"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
list.Clear

i = 0
Do While Not Data1.Recordset.EOF
    If Data1.Recordset.EOF Then Exit Do
    menus(i) = Data1.Recordset!nome & "||" & Data1.Recordset!Path & "||" & Data1.Recordset!opcoes
    i = i + 1
    Data1.Recordset.MoveNext
Loop
For i = 0 To UBound(menus)
    If menus(i) <> "" Then
        list.AddItem Mid(menus(i), 1, InStr(menus(i), "||") - 1)
    End If
Next
salvou = "sim"
recarregar
list.ListIndex = 0
salvar = "sim"
End Sub
Private Sub recarregar()
Text1 = ""
Text2 = ""
Text3 = ""
antes = list.ListIndex
list.Clear
For i = 0 To UBound(menus)
    If menus(i) <> "" Then
        list.AddItem Mid(menus(i), 1, InStr(menus(i), "||") - 1)
    End If
Next
list.ListIndex = 0
Command5.Caption = "Atualizar"
Command8.Visible = True
Command1.Visible = True
Command7.Visible = False
Command11.Visible = False
salvar = "nao"
End Sub

Private Sub separar(envstr As String)
If envstr = "" Then Exit Sub
temp = envstr
nome = Mid(temp, 1, InStr(1, temp, "||") - 1)
temp = Mid(temp, Len(nome) + 3, Len(temp))
Path = Mid(temp, 1, InStr(1, temp, "||") - 1)
temp = Mid(temp, Len(Path) + 3, Len(temp))
opcoes = temp
mntemp(0) = nome
If nome = "-" Then
Path = "[ SEPARADOR ]"
End If

mntemp(1) = Path
mntemp(2) = opcoes
End Sub


Private Sub list_Click()
separar menus(list.ListIndex)
Text1 = mntemp(0)
Text2 = mntemp(1)
Text3 = mntemp(2)
Command8.Visible = True
Command5.Caption = "Atualizar"
Command1.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command5_Click
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command5_Click
End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command5_Click
End If
End Sub
