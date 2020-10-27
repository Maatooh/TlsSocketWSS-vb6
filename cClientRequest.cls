VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClientRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID                   As Long
Public WithEvents Socket    As cTlsSocket
Attribute Socket.VB_VarHelpID = -1
Public Callback             As Object
Public HasHandshake         As Boolean

Private Sub Class_Initialize()
    HasHandshake = False
End Sub

Private Sub Socket_OnReceive()
    Callback.ClientOnReceive Me
End Sub

Private Sub Socket_OnClose()
    HasHandshake = False
    Callback.ClientOnClose Me
End Sub
