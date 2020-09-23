VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   Caption         =   "NoIcons - Opções"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   4785
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000007&
      Caption         =   "Adicionar NoIcons no registro para rodar automaticamente"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Checked = True Then
    Call savestring("hkey_local_machine", "Software\microsoft\windows\currentversion\run", "NoIcons", App.Path & "\" & App.EXEName & ".exe")
Else
    Call DeleteValue("hkey_local_machine", "Software\microsoft\windows\currentversion\run", "NoIcons")
End Sub

Private Sub Form_Load()
X = getdword("hkey_local_machine", "Software\microsoft\windows\currentversion\run", "NoIcons")
If X = "" Then
    Check1.Checked = False
Else
    Check1.Checked = True
End If

    
End Sub
