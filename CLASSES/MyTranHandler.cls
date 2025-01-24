VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyTranHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements PosInterface.ITransactionDoneHandler

Public Sub ITransactionDoneHandler_OnFinish(ByVal message As String)

    MsgBox (message)
    TransactionFinished
End Sub

Public Sub ITransactionDoneHandler_OnTransactionDone(ByVal Result As PosInterface.TransactionResult)
 'MsgBox (result.ToString)

    TransactionDone Result

End Sub

