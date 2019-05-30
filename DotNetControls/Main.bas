Attribute VB_Name = "Module1"
Public Sub Main()
    Dim oInitializationService As InitializationService
    Set oInitializationService = New InitializationService
    oInitializationService.Initialize
    
    Dim oForm As Form1
    Set oForm = New Form1
    oForm.Show
End Sub
