VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "agenda.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin NoIcons43.Button3D Command1 
      Height          =   330
      Left            =   7065
      TabIndex        =   2
      Top             =   3285
      Width           =   1185
      _ExtentX        =   2090
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
   Begin VB.ComboBox cmes 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   450
      TabIndex        =   1
      Top             =   3285
      Width           =   1725
   End
   Begin MSDBGrid.DBGrid grid 
      Bindings        =   "agenda.frx":058A
      Height          =   3255
      Left            =   45
      OleObjectBlob   =   "agenda.frx":059A
      TabIndex        =   0
      Top             =   0
      Width           =   8205
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7155
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2565
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mes(1 To 12) As Variant


Private Sub cmes_Click()
If cmes.text = "Janeiro" Then X = 1
If cmes.text = "Fevereiro" Then X = 2
If cmes.text = "Março" Then X = 3
If cmes.text = "Abril" Then X = 4
If cmes.text = "Maio" Then X = 5
If cmes.text = "Junho" Then X = 6
If cmes.text = "Julho" Then X = 7
If cmes.text = "Agosto" Then X = 8
If cmes.text = "Setembro" Then X = 9
If cmes.text = "Outubro" Then X = 10
If cmes.text = "Novembro" Then X = 11
If cmes.text = "Dezembro" Then X = 12
grid.Caption = "Agenda para o mês de " & cmes.text
Data1.RecordSource = "select * from agenda where data like '*/0" & X & "/*' order by data"
Data1.Refresh
grid.Refresh

End Sub

Private Sub cmes_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMM
Me.Icon = Form1.img.ListImages(2).Picture
Me.Caption = versao & " - Agenda"
Data1.DatabaseName = App.Path & "\noicons.mdb"

mes(1) = "Janeiro"
mes(2) = "Fevereiro"
mes(3) = "Março"
mes(4) = "Abril"
mes(5) = "Maio"
mes(6) = "Junho"
mes(7) = "Julho"
mes(8) = "Agosto"
mes(9) = "Setembro"
mes(10) = "Outubro"
mes(11) = "Novembro"
mes(12) = "Dezembro"
For i = 1 To 12
    If i = CInt(Right(Format(Date, "mm"), Len(i))) Then
        X = i
    End If
    
    cmes.AddItem mes(i)
Next
cmes.ListIndex = X - 1
cmes_Click

End Sub


Private Sub grid_AfterColUpdate(ByVal ColIndex As Integer)
Data1.Refresh
grid.Refresh
End Sub


Private Sub grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
On Error Resume Next

End Sub
