VERSION 5.00
Begin VB.UserControl Bevel 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   ScaleHeight     =   4050
   ScaleWidth      =   4725
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00808080&
      Height          =   15
      Left            =   240
      ScaleHeight     =   15
      ScaleWidth      =   4455
      TabIndex        =   3
      Top             =   3690
      Width           =   4455
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00808080&
      Height          =   3165
      Left            =   4560
      ScaleHeight     =   3165
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   675
      Width           =   15
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2925
      Left            =   390
      ScaleHeight     =   2925
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   840
      Width           =   15
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   255
      ScaleHeight     =   15
      ScaleWidth      =   4470
      TabIndex        =   0
      Top             =   795
      Width           =   4470
   End
End
Attribute VB_Name = "Bevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Enum TStyle
  VbLowered = 0
  VbRaised = 2
End Enum

Sub BevelSytle(Style As TStyle)
    Select Case Style
        Case VbLowered
            Picture1.BackColor = &H808080
            Picture2.BackColor = &H808080
            Picture3.BackColor = &HFFFFFF
            Picture4.BackColor = &HFFFFFF
            
        Case VbRaised
            Picture1.BackColor = &HFFFFFF
            Picture2.BackColor = &HFFFFFF
            Picture3.BackColor = &HFFFFFF
            Picture4.BackColor = &HFFFFFF
            
    End Select
    
End Sub

Private Sub UserControl_Resize()
Picture1.Move 0, 0, 20, UserControl.Height
Picture2.Move 0, 0, UserControl.Width, 20
Picture3.Move UserControl.ScaleWidth - 20, 0, 20, UserControl.Height
Picture4.Move 0, UserControl.ScaleHeight - 20, UserControl.Width, 20

End Sub

