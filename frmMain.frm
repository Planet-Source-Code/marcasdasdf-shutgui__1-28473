VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShutGUI"
   ClientHeight    =   4575
   ClientLeft      =   4245
   ClientTop       =   2640
   ClientWidth     =   4230
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4230
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   350
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   350
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "Shutdown"
      Default         =   -1  'True
      Height          =   350
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtDelay 
      Height          =   300
      Left            =   2490
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "30"
      Top             =   3550
      Width           =   465
   End
   Begin VB.TextBox txtMessage 
      Height          =   975
      Left            =   120
      MaxLength       =   127
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   3975
      Begin VB.CheckBox chkReboot 
         Caption         =   "Reboot After Shutdown"
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   1960
      End
      Begin VB.CheckBox chkKillApps 
         Caption         =   "Kill Applications Without Saving Data"
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   360
         Width           =   2920
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   300
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   200
      Width           =   375
   End
   Begin VB.TextBox txtComputer 
      Height          =   300
      Left            =   1410
      TabIndex        =   0
      Top             =   200
      Width           =   2175
   End
   Begin VB.Label lblDelay 
      Caption         =   "Delay (seconds)"
      Height          =   255
      Left            =   1215
      TabIndex        =   12
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label lblMessage 
      Caption         =   "Message to Send (127 characters max)"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   2795
   End
   Begin VB.Label lblComputer 
      AutoSize        =   -1  'True
      Caption         =   "Computer Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   1170
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbort_Click()
    If txtComputer.Text = "" Then
        MsgBox "You must specify a computer before clicking the Abort button.  ", vbExclamation + vbOKOnly, "Specify a Computer"
    Else
        cancelShutdown = AbortShutdown(txtComputer.Text)
        If cancelShutdown = 1 Then
            MsgBox "Successfully aborted the shutdown of  '" & txtComputer.Text & "'.  ", vbInformation + vbOKOnly, "Abort Successful"
        ElseIf cancelShutdown = 0 Then
            MsgBox "Unable to abort the shutdown of  '" & txtComputer.Text & "'.  Reasons may include:  " & vbCrLf & vbCrLf & "1. You misspelled the computer name." & vbCrLf & "2. The computer is not a Window NT computer." & vbCrLf & "3. You do not have access to reboot the computer." & vbCrLf & "4. The computer cannot be found on the network." & vbCrLf & "5. It's too late.", vbExclamation + vbOKOnly, "Cannot Abort"
        End If
    End If
End Sub

Private Sub cmdBrowse_Click()
    txtComputer.Text = GetBrowseNetworkWorkstation()
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdShutdown_Click()
    If txtComputer.Text = "" Then
        MsgBox "You must specify a computer before clicking the Shutdown button.  ", vbExclamation + vbOKOnly, "Specify a Computer"
    Else
        goShutdown = ShutDownComputer(txtComputer.Text, txtMessage.Text, txtDelay.Text, chkKillApps.Value, chkReboot.Value)
        If goShutdown = 1 Then
            MsgBox "Successfully sent a shutdown request to  '" & txtComputer.Text & "'.  " & vbCrLf & "The computer will shutdown in " & txtDelay.Text & " seconds.  ", vbInformation + vbOKOnly, "Shutdown Sent"
        ElseIf goShutdown = 0 Then
            MsgBox "Unable to shutdown  '" & txtComputer.Text & "'.  Reasons may include:  " & vbCrLf & vbCrLf & "1. You misspelled the computer name." & vbCrLf & "2. The computer is not a Window NT computer." & vbCrLf & "3. You do not have access to reboot the computer." & vbCrLf & "4. The computer cannot be found on the network.", vbExclamation + vbOKOnly, "Cannot Shutdown"
        End If
    End If
End Sub

Private Sub Form_Load()
    makeCppButtons Me
    centerForm Me
    txtComputer.Text = GetSetting("ShutGUI", "Settings", "ComputerName", "")
    txtMessage.Text = GetSetting("ShutGUI", "Settings", "Message", "")
    txtDelay.Text = GetSetting("ShutGUI", "Settings", "Delay", "30")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ShutGUI", "Settings", "ComputerName", txtComputer.Text
    SaveSetting "ShutGUI", "Settings", "Message", txtMessage.Text
    SaveSetting "ShutGUI", "Settings", "Delay", txtDelay.Text
End Sub

