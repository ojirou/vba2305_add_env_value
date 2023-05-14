Attribute VB_Name = "Module1"
'#############################################################################
' ƒ†[ƒU[ŠÂ‹«•Ï”‚ğ’Ç‰Á
'
'
'#############################################################################
Option Explicit
Sub add_env_value(EnvName, EnvValue As String)
Dim idx, envString
idx = 1
Do
    envString = Environ(idx)
    idx = idx + 1
    Debug.Print envString
Loop Until envString = ""
Dim wshShell, wshUserEnv
Set wshShell = CreateObject("WScript.Shell")
Set wshUserEnv = wshShell.Environment("User")
wshUserEnv.Item(EnvName) = EnvValue
Set wshUserEnv = Nothing
Set wshShell = Nothing
End Sub
