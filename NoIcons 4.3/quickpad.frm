VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form quickpad 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "quickpad.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin NoIcons43.Button3D Button3D1 
      Height          =   375
      Left            =   7155
      TabIndex        =   3
      Top             =   3735
      Width           =   1050
      _ExtentX        =   1852
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
   Begin NoIcons43.Button3D Command1 
      Height          =   375
      Left            =   6030
      TabIndex        =   2
      Top             =   3735
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      BackColor       =   0
      Caption         =   "Exportar"
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
      Height          =   375
      Left            =   4905
      TabIndex        =   1
      Top             =   3735
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      BackColor       =   0
      Caption         =   "Salvar"
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
   Begin RichTextLib.RichTextBox texto 
      Height          =   3480
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   6138
      _Version        =   393217
      BackColor       =   -2147483639
      ScrollBars      =   3
      TextRTF         =   $"quickpad.frx":0442
   End
End
Attribute VB_Name = "quickpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button3D1_Click()
If qsalvou <> "sim" Then
    x = MsgBox("Você fez alterações no texto e ainda não salvou." & Chr(10) & "Deseja salvar agora?", vbYesNo + vbQuestion, "QuickPad")
    If x = vbYes Then
            Command2_Click
    End If
End If

Unload Me
    
End Sub

Private Sub Command1_Click()
Dim sSave As SelectedFile
Dim Count As Integer
Dim FileList As String

    On Error GoTo error
    
    FileDialog.sFilter = "Text Files (*.txt)" & Chr$(0) & "*.txt"
    
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sDlgTitle = "Exportar"
    FileDialog.sInitDir = App.Path & "\"
    sSave = ShowSave(Me.hWnd)
    If Err.Number <> 32755 And sSave.bCanceled = False Then
    sFile = sSave.sLastDirectory
    If InStr(1, LCase(sSave.sFiles(1)), ".txt") = 0 Then
        sFile = sFile & sSave.sFiles(1) & ".txt"
    Else
        sFile = sFile & sSave.sFiles(1)
    End If
    If Dir(sFile) <> "" Then
        x = MsgBox("O arquivo já existe, deseja escrever por cima?", vbYesNo + vbExclamation, "QuickPad")
        If x = vbNo Then Exit Sub
        Kill sFile
    End If
    Open sFile For Output As #1
        Print #1, texto.text
    Close #1
    x = MsgBox("O Arquivo foi exportado com sucesso!" & Chr(10) & "Deseja abrir ele agora?", vbYesNo + vbQuestion, "QuickPad")
    If x = vbYes Then Executar sFile, "", ""
    End If
    Exit Sub
error:
Debug.Print Err.Number & ": " & Err.Description
Exit Sub
End Sub

Private Sub Command2_Click()
Open App.Path & "\quickpad.dat" For Output As #1
Print #1, texto.text
Close #1
qsalvou = "sim"
End Sub

Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMM
Me.Caption = "QuickPad - " & versao
If Dir(App.Path & "\quickpad.dat") = "" Then
Open App.Path & "\quickpad.dat" For Append As #1
Print #1, "NoIcons 4.3 Build 04/06/01"
Close #1
End If
texto.LoadFile App.Path & "\quickpad.dat"
qsalvou = "sim"
End Sub
Private Sub texto_KeyPress(KeyAscii As Integer)
qsalvou = "nao"
End Sub
