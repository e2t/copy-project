Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object

Sub Main()
    Dim doc As ModelDoc2
    Dim oldStateUseFolderSearchRules As Boolean
    
    Set swApp = Application.SldWorks
    Set doc = swApp.ActiveDoc
    If doc Is Nothing Then Exit Sub
    oldStateUseFolderSearchRules = swApp.GetUserPreferenceToggle(swUseFolderSearchRules)
    swApp.SetUserPreferenceToggle swUseFolderSearchRules, False
    swApp.RunCommand swCommands_File_Copy_Design, Empty
    swApp.SetUserPreferenceToggle swUseFolderSearchRules, oldStateUseFolderSearchRules
End Sub
