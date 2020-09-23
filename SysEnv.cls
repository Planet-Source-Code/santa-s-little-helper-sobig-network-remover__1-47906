VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sysENV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Base 0
Option Explicit
Public SysEnvObj As New WshShell

Public Function GetEnv(VarName As String) As String
On Local Error Resume Next
    GetEnv = SysEnvObj.Environment.Item(VarName)
On Local Error GoTo 0
End Function

Public Function SetEnv(VarName As String, VarVal As String) As String
On Local Error Resume Next
    SysEnvObj.Environment.Item(VarName) = VarVal
    SetEnv = SysEnvObj.Environment.Item(VarName)
On Local Error GoTo 0
End Function

Public Function DelEnv(VarName As String) As String
On Local Error Resume Next
    SysEnvObj.Environment.Remove (VarName)
    DelEnv = "Deleted"
On Local Error GoTo 0
End Function
