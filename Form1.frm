VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "...."
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Subclass"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Unsubclass"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      _cx             =   4206819
      _cy             =   4201315
      Movie           =   "catMoreMail.swf"
      Src             =   "catMoreMail.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu mnuMy 
         Caption         =   "&My"
      End
      Begin VB.Menu mnuCustom 
         Caption         =   "&Custom"
      End
      Begin VB.Menu mnuMenu 
         Caption         =   "&Menu"
      End
      Begin VB.Menu mnuFor 
         Caption         =   "&For"
      End
      Begin VB.Menu mnuFlash 
         Caption         =   "&Flash"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@@@@@@@@@ Developed by Ravindra Deuskar @@@@@@@@@@@@@@@@@@@@
Option Explicit

Private Sub cmdExit_Click()
    If FHW <> 0 Then UnSubClass
    MsgBox "Please vote for me"
    End
End Sub

Private Sub Command1_Click()
Dim lRet As Long, lParam As Long
Dim lhWnd As Long

lhWnd = Me.hwnd
lRet = EnumChildWindows(lhWnd, AddressOf EnumChildProc, lParam)

If FHW <> 0 Then
   glPrevWndProc = SubClass()
End If
End Sub

Private Sub Command2_Click()
    If FHW <> 0 Then UnSubClass
End Sub

Private Sub Form_Load()
    ShockwaveFlash1.WMode = "Window" ' Important B'cause you want hWnd of flash activex control
    ShockwaveFlash1.Movie = App.Path + "\Dispablerightclick.swf"
End Sub
