VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmIcon 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3225
   ClientLeft      =   1410
   ClientTop       =   1530
   ClientWidth     =   8805
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   8805
   Begin NoIcons43.Button3D Command3 
      Height          =   285
      Left            =   2745
      TabIndex        =   5
      Top             =   360
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   503
      BackColor       =   0
      Caption         =   "Procurar"
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
      Height          =   285
      Left            =   2745
      TabIndex        =   4
      Top             =   45
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   503
      BackColor       =   0
      Caption         =   "Default"
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
   Begin VB.Data Data2 
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
      Top             =   3105
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox atual 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   1800
      ScaleHeight     =   555
      ScaleWidth      =   600
      TabIndex        =   3
      Top             =   45
      Width           =   600
   End
   Begin MSComctlLib.ImageList icons 
      Left            =   7785
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":11FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":1516
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":1832
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":1B4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":1E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":2186
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":24A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":27BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":2ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":2DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":3112
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":342E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":374A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":3A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":3D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":409E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":43BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":46D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":49F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":4D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":502A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":5346
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":5662
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":597E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":5C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":5FB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":62D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":65EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":690A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":6C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":6F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":725E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":757A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":7896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picon 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   0
      Left            =   90
      ScaleHeight     =   555
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   855
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   150
      Left            =   7380
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atual:"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   180
      Width           =   600
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
atual.Picture = Form1.img.ListImages(1).Picture
Data2.Recordset.Edit
Data2.Recordset!current = "Default"
Data2.Recordset!Path = ""
Data2.Recordset.Update
Form1.SysIcon.IconHandle = Form1.img.ListImages(1).Picture

End Sub

Private Sub Command1_Click()
' 1 até 14
For i = 1 To 14
Load picon(i)
If i = 1 Then
    picon(i).left = picon(i - 1).left
Else
    picon(i).left = picon(i - 1).left + 600
End If
    picon(i).top = 855
    picon(i).Visible = True
    picon(i).Picture = icons.ListImages(i).Picture
Next
'15 até 28
For i = 15 To 28
Load picon(i)
If i = 15 Then
    picon(i).left = picon(0).left
Else
    picon(i).left = picon(i - 1).left + 600
End If
    picon(i).top = 1440
    picon(i).Visible = True
    picon(i).Picture = icons.ListImages(i).Picture
Next

'29 até 38
For i = 29 To icons.ListImages.Count - 1
Load picon(i)
If i = 29 Then
    picon(i).left = picon(0).left
Else
    picon(i).left = picon(i - 1).left + 600
End If
    picon(i).top = 2025
    picon(i).Visible = True
    picon(i).Picture = icons.ListImages(i).Picture
Next
End Sub

Private Sub Command3_Click()
Dim sOpen As SelectedFile
Dim Count As Integer
Dim FileList As String

    On Error GoTo error
    FileDialog.sFilter = "ïcones (*.ico)" & Chr$(0) & "*.ico"
    FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
    FileDialog.sDlgTitle = "Abrir Arquivo"
    FileDialog.sInitDir = App.Path & "\"
    sOpen = ShowOpen(Me.hWnd)
    spath = sOpen.sLastDirectory & sOpen.sFiles(1)
    If Dir(App.Path & "\icon.ni4") <> "" Then Kill App.Path & "\icon.ni4"
    FileCopy spath, App.Path & "\icon.ni4"
    atual.Picture = LoadPicture(App.Path & "\icon.ni4")
    Data2.Recordset.Edit
    Data2.Recordset!current = "user"
    Data2.Recordset!Path = App.Path & "\icon.ni4"
    Data2.Recordset.Update
    Form1.SysIcon.IconHandle = LoadPicture(App.Path & "\icon.ni4")

error:
    Exit Sub

End Sub

Private Sub Form_Load()
Me.Caption = versao & " - Ícones"
Data2.DatabaseName = App.Path & "\noicons.mdb"
Data2.RecordSource = "select * from icon"
Data2.Refresh
If Data2.Recordset!current = "Default" Then
    atual.Picture = Form1.img.ListImages(1).Picture
Else
    atual.Picture = LoadPicture(Data2.Recordset!Path)
End If
Command1_Click

End Sub

Private Sub picon_Click(Index As Integer)
If Dir(App.Path & "\icon.ni4") <> "" Then Kill App.Path & "\icon.ni4"
SavePicture picon(Index).Picture, App.Path & "\icon.ni4"
    
Data2.Recordset.Edit
Data2.Recordset!current = "user"
Data2.Recordset!Path = App.Path & "\icon.ni4"
Data2.Recordset.Update
Form1.SysIcon.IconHandle = LoadPicture(App.Path & "\icon.ni4")
atual.Picture = LoadPicture(App.Path & "\icon.ni4")
End Sub
