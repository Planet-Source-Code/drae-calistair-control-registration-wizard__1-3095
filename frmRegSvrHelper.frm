VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRegSvrHelper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Registration"
   ClientHeight    =   840
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   1800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   688
      Style           =   1
      SimpleText      =   "Status: Waiting:"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileRegister 
         Caption         =   "&Register Control"
      End
      Begin VB.Menu mnuFileUnregister 
         Caption         =   "&UnRegister Control"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmRegSvrHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'****************************************************
'* Control Registration Wizard                      *
'* by Drae Calistair                                *
'*                                                  *
'* Uses: To register and unregister OCX and DLL     *
'*       files for use with Visual Basic 5.0 and    *
'*       other programming languages that support   *
'*       OCXs and DLLs.                             *
'*                                                  *
'* Drae Calistair makes no guarentee on the         *
'* functionality or results of this code. As with   *
'* all other code fragments, use at your own risk.  *
'* Drae Calistair is not responsible to damages     *
'* caused by misuse, intentional or otherwise, of   *
'* this source code.                                *
'****************************************************
End Sub

Private Sub mnuFileExit_Click()
' Unloads the form and ends the program.
Unload Me
End
End Sub

Private Sub mnuFileRegister_Click()
' Allows user to specify an OCX or DLL file to register

stbStatus.SimpleText = "Status: Choose a file to register:"

' Sets up the Common Dialog control to retrieve the file-
' name to register.
cmd1.filename = ""
cmd1.DialogTitle = "Choose a file to register:"
cmd1.Filter = "*.OCX (Active-X Control)|*.ocx|*.DLL (Dynamic Link Library file)|*.dll"
cmd1.ShowOpen

' Checks to see if user canceled. If so, updates the
' status bar text and canceles the process.
If cmd1.filename = vbNullString Or cmd1.FileTitle = vbNullString Then
    stbStatus.SimpleText = "Status: Waiting:"
    Exit Sub
End If

' Calls the control registration program and specifies
' the control or dll to register (using Win98 or Win95).
a = Shell("C:\Windows\System\regsvr32 " + cmd1.filename, vbNormalNoFocus)

stbStatus.SimpleText = "Status: Registering " + cmd1.filename

' Tries again if the user is using WinNT.
If a = 0 Then
    b = Shell("C:\WinNT\System\regsvr32 " + cmd1.filename, vbHide)
End If

stbStatus.SimpleText = "Status: Waiting:"

End Sub

Private Sub mnuFileUnregister_Click()
' Allows user to specify an OCX or DLL file to unregister

stbStatus.SimpleText = "Status: Choose a file to unregister:"

' Sets up the Common Dialog control to retrieve the file-
' name to register.
cmd1.filename = ""
cmd1.DialogTitle = "Choose a file to unregister:"
cmd1.Filter = "*.OCX (Active-X Control)|*.ocx"
cmd1.ShowOpen

' Checks to see if user canceled. If so, updates the
' status bar text and canceles the process.
If cmd1.filename = vbNullString Or cmd1.FileTitle = vbNullString Then
    stbStatus.SimpleText = "Status: Waiting:"
    Exit Sub
End If

' Calls the control registration program and specifies
' the control or dll to register (using Win98 or Win95).
a = Shell("C:\Windows\System\regsvr32 /u " + cmd1.filename, vbNormalNoFocus)

stbStatus.SimpleText = "Status: Unregistering " + cmd1.filename

' Tries again if the user is using WinNT.
If a = 0 Then
    b = Shell("C:\WinNT\System\regsvr32 /u " + cmd1.filename, vbHide)
End If

stbStatus.SimpleText = "Status: Waiting:"

End Sub
