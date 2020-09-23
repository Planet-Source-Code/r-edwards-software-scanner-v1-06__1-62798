Attribute VB_Name = "Env_Var"
'------------------------------------------------------------
' Module    : modEnv_Var.bas
' Purpose   : Get PC Info
' Author    : unknown
' Source    : Planet Source Code
'------------------------------------------------------------
'
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'------------------------------------------------------------
' Purpose   : Returns the name of the computer
' Arguments : none
'------------------------------------------------------------
'
Public Function ComputerName() As String
    
    Dim buffer As String * 512
    Dim length As Long
    
    length = Len(buffer)
    If GetComputerName(buffer, length) Then
        ' this API returns non-zero if successful,
        ' and modifies the length argument
        ComputerName = Left$(buffer, length)
    End If

End Function

