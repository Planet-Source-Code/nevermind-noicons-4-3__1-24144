VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form InfoHd 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   3690
   ClientTop       =   2430
   ClientWidth     =   4305
   Icon            =   "InfoHd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4305
   Begin MSComctlLib.ProgressBar free1 
      Height          =   375
      Left            =   405
      TabIndex        =   3
      Top             =   2610
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   405
      TabIndex        =   0
      Top             =   45
      Width           =   2400
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizado:"
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
      Left            =   405
      TabIndex        =   4
      Top             =   2340
      Width           =   990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1905
      Left            =   405
      TabIndex        =   2
      Top             =   405
      Width           =   3840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   3465
      TabIndex        =   1
      Top             =   2610
      Width           =   585
   End
End
Attribute VB_Name = "InfoHd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ddd As New clsFuncoes

Private Sub Drive1_Change()
checka (Mid(Drive1.Drive, 1, 1))
End Sub

Private Sub Form_Load()
Me.Caption = versao & " - Informações do HD"
free1.Value = 0
checka (Mid(Drive1.Drive, 1, 1))
End Sub
Private Sub checka(Drive As String)
x = "Tipo: " & ddd.Get_DriveType(Drive) & Chr(10)
x = x & ddd.FreeDiscSpace(Drive)
Label2 = x
End Sub
