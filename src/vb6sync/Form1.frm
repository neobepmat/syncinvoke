VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textCounter 
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox textLog 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0005
      Top             =   120
      Width           =   4335
   End
   Begin VB.Timer timerLoop 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   3120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cSyncLib As cSyncLibCOM
Attribute cSyncLib.VB_VarHelpID = -1
Private msgFromSyncLib As String

Private Sub cSyncLib_OnNewEventFromSyncLib(msg As String)
Call loopRoutine(msg)
End Sub

Private Sub Form_Load()
Set cSyncLib = New cSyncLibCOM
End Sub

Private Sub loopRoutine(msg As String)
Dim i As Long
Dim mymsg As String

textLog.Text = "Incoming Message: " + msg + vbCr + vbLf

mymsg = "Loop started at: " + CStr(Now())
textLog.Text = textLog + mymsg + vbCr + vbLf
For i = 0 To 50000
    textCounter = CStr(i)
    DoEvents
Next i

mymsg = "Loop Finished at: " + CStr(Now())
textLog.Text = textLog + mymsg + vbCr + vbLf

timerLoop.Enabled = False
End Sub
