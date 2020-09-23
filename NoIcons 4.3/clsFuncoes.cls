VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const MAX_TIP_LENGTH As Long = 64
Private nidTrayIcon As NOTIFYICONDATA
Private bIconDisplayed, bUpdateOnChange As Boolean
Public Event NIError(ByVal ErrorNumber As Long)
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TIP_LENGTH
End Type


Public Sub download(sUrl As String, sFile As String)
' I set it up so the url is different then the file so all you have to do is set up the url once then just change the File Name
Dim DL As Long
On Error GoTo errHandler
If sUrl$ = "" Then sUrl$ = strUrl$
If strUrl$ = "" Then strUrl$ = sUrl$
If Left(strUrl$, 4) <> "http" Then strUrl$ = "http://" & strUrl$ 'See if  there's a http:// at the begining if not add it
If Right$(strUrl$, 1) <> "/" Then strUrl$ = strUrl$ & "/" 'see if there is an "/" at the end of the url if theres not add it
If Left$(sFile$, 1) = "/" Then sFile = Mid$(sFile$, 2) 'since we added an "/" we have to make sure there isnt one at the front of the file"
 DL& = DoFileDownload(StrConv(strUrl$ & sFile$, vbUnicode))
 Exit Sub
errHandler:
Debug.Print "Error Source:", Err.Source
Debug.Print "Error Description:", Err.Description
Debug.Print "Error Number:", Err.Number
MsgBox "Ocorreu um erro ao tentar fazer o download do arquivo " & sUrl & sFile$ & ".", vbOKOnly + vbExclamation, "Live Update - Erro"
End Sub

Public Function FreeDiscSpace(sDriveLetter As String) As String
    Dim Sectors As Long, Bytes As Long, FreeC As Long, TotalC As Long, Total As Long, Freeb As Long, sTemp As String
    'Retrieve information about the sDriveLetter
    On Error GoTo erro
    InfoHd.free1.Visible = True
    InfoHd.Label3.Visible = True
    GetDiskFreeSpace sDriveLetter$ & ":\", Sectors, Bytes, FreeC, TotalC
    sTemp$ = "Setores por Cluster:" & Str$(Sectors) & vbCrLf
    sTemp$ = sTemp$ & "Bytes por Setor:" & Str$(Bytes) & vbCrLf
    sTemp$ = sTemp$ & "Total de Clusters:" & Str$(TotalC) & vbCrLf
    sTemp$ = sTemp$ & "Clusters Livres:" & Str$(FreeC) & vbCrLf
    Total = TotalC& * Sectors& * Bytes&
    sTemp$ = sTemp$ & "Espaço Total: " & Mid(Str$(Total) / 1048576, 1, InStr(1, Str$(Total) / 1048576, ",") + 2) & " MB" & vbCrLf
    InfoHd.free1.Max = Str$(Total) / 1048576
    Freeb = FreeC& * Sectors& * Bytes&
    sTemp$ = sTemp$ & "Espaço Livre: " & Mid(Str$(Freeb) / 1048576, 1, InStr(1, Str$(Freeb) / 1048576, ",") + 2) & " MB" & vbCrLf
    InfoHd.free1.Value = InfoHd.free1.Max - Str$(Freeb) / 1048576
    FreeDiscSpace$ = sTemp$
    InfoHd.Label1.Caption = Mid(InfoHd.free1.Value * 100 / InfoHd.free1.Max, 1, InStr(1, InfoHd.free1.Value * 100 / InfoHd.free1.Max, ",") + 1) & "%"
erro:
    If Err.Number <> 0 Then
    Debug.Print Err.Number & ": " & Err.Description
    InfoHd.free1.Value = 0
    InfoHd.Label1.Caption = ""
    InfoHd.free1.Visible = False
    InfoHd.Label3.Visible = False
    End If
End Function

Public Function Get_DriveType(sDriveLetter As String) As String
Select Case GetDriveType(sDriveLetter$ & ":\")
        Case 2
            Get_DriveType$ = "Removível"
        Case 3
            Get_DriveType$ = "Drive Fixo"
        Case Is = 4
            Get_DriveType$ = "Remoto"
        Case Is = 5
            Get_DriveType$ = "CD-Rom"
        Case Is = 6
            Get_DriveType$ = "Ram disk"
        Case Else
            Get_DriveType$ = "Desconhecido"
    End Select
End Function

Public Sub MakeNormal(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Function Initialize(ByVal hWnd As Long, ByVal hIcon As Long, ByVal sTip As String, Optional ByVal uCallbackMessage As Long = WM_MOUSEMOVE) As Long
  With nidTrayIcon
   .cbSize = Len(nidTrayIcon)
   .hIcon = hIcon
   .hWnd = hWnd
   .szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
   .uCallbackMessage = uCallbackMessage
   .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
   .uId = vbNull
  End With
  bIconDisplayed = False
  bUpdateOnChange = True
End Function

Public Function ShowIcon() As Boolean
  If Not bIconDisplayed Then
    ShowIcon = Shell_NotifyIcon(NIM_ADD, nidTrayIcon)
    If ShowIcon = False Then
      RaiseEvent NIError(GetLastError)
    Else
      bIconDisplayed = True
    End If
  End If
End Function

Public Function HideIcon() As Boolean
  If bIconDisplayed Then
    HideIcon = Shell_NotifyIcon(NIM_DELETE, nidTrayIcon)
    If HideIcon = False Then
      RaiseEvent NIError(GetLastError)
    Else
      bIconDisplayed = False
    End If
  End If
End Function

Public Property Let IconHandle(ByVal hIcon As Long)
  nidTrayIcon.hIcon = hIcon
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_ICON
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Property Let TipText(ByVal sTip As String)
  nidTrayIcon.szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_TIP
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Property Let CallbackMessage(ByVal uCallbackMessage As Long)
  nidTrayIcon.uCallbackMessage = uCallbackMessage
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_MESSAGE
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property

Public Function Update() As Boolean
  If bIconDisplayed Then
    Update = Shell_NotifyIcon(NIM_MODIFY, nidTrayIcon)
    If Update = False Then
      RaiseEvent NIError(GetLastError)
    End If
  End If
End Function

Public Property Get IconHandle() As Long
  IconHandle = nidTrayIcon.hIcon
End Property

Public Property Get TipText() As String
  TipText = Left(nidTrayIcon.szTip, Len(nidTrayIcon.szTip) - 1)
End Property

Public Property Get CallbackMessage() As Long
  CallbackMessage = nidTrayIcon.uCallbackMessage
End Property

Public Property Let UpdateOnChange(bUpdate As Boolean)
  bUpdateOnChange = bUpdate
End Property

Private Property Get UpdateOnChange() As Boolean
  UpdateOnChange = bUpdateOnChange
End Property

Private Sub Class_Terminate()
  HideIcon
End Sub

Public Property Get Visible() As Boolean
  If bIconDisplayed Then
    Visible = True
  End If
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
  If bVisible Then
    ShowIcon
  Else
    HideIcon
  End If
End Property

