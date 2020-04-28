VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSyncLibCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oSyncLibCom As clibSync.SyncLibCOM
Attribute oSyncLibCom.VB_VarHelpID = -1
Public Event OnNewEventFromSyncLib(msg As String)

Private Sub Class_Initialize()
Set oSyncLibCom = New clibSync.SyncLibCOM
oSyncLibCom.Init
End Sub

Private Sub Class_Terminate()
Set oSyncLibCom = Nothing
End Sub

Private Sub oSyncLibCom_OnSyncLibEvent(ByVal msg As String)
RaiseEvent OnNewEventFromSyncLib(msg)
End Sub
