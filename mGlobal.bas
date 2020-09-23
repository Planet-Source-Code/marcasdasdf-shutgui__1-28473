Attribute VB_Name = "mGlobal"
Option Explicit

Private Const ERROR_SUCCESS As Long = 0
Private Const MAX_PATH As Long = 260
Private Const CSIDL_NETWORK As Long = &H12
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000

Private Type BROWSEINFO
   hOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Integer, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function ShutDownComputer(CompName As String, MessageToUser As String, SecondsUntilShutdown As Long, ForceAppsClosed As Integer, RebootAfter As Integer)
    ShutDownComputer = InitiateSystemShutdown(CompName, MessageToUser, SecondsUntilShutdown, ForceAppsClosed, RebootAfter)
End Function

Public Function AbortShutdown(CompName As String)
     AbortShutdown = AbortSystemShutdown(CompName)
End Function

Public Function GetBrowseNetworkWorkstation() As String
    Dim BI As BROWSEINFO
    Dim pidl As Long
    Dim sPath As String
    Dim pos As Integer
    If SHGetSpecialFolderLocation(frmMain.hWnd, CSIDL_NETWORK, pidl) = ERROR_SUCCESS Then
        With BI
            .hOwner = frmMain.hWnd
            .pidlRoot = pidl
            .pszDisplayName = Space$(MAX_PATH)
            .lpszTitle = "Select a network computer."
            .ulFlags = BIF_BROWSEFORCOMPUTER
        End With
        If SHBrowseForFolder(BI) <> 0 Then
            GetBrowseNetworkWorkstation = "\\" & BI.pszDisplayName
        End If
        Call CoTaskMemFree(pidl)
    End If
End Function

Public Function makeCppButtons(frm As Form) As Long
    Const BM_SETSTYLE = &HF4
    Const BS_SOLID = 0
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is CommandButton Then
            If ctrl.Style = 1 Then
                SendMessage ctrl.hWnd, BM_SETSTYLE, BS_SOLID, 1
            'Else
                'MsgBox "The style for the button  '" & ctrl.Name & "'  on form  '" & frm.Name & "'  is not set" & vbCrLf & "to  '1 - Graphical'. The button will be shown as a normal VB button.", vbInformation + vbOKOnly, "Can't Make C++ Button"
            End If
        End If
    Next
End Function

Public Function centerForm(frm As Form) As Long
    Dim winLeft As Integer
    Dim winTop As Integer
    winLeft = ((Screen.Width - frm.Width) / 2) + 60
    winTop = ((Screen.Height - 800) - frm.Height) / 2
    frm.Move winLeft, winTop
End Function

