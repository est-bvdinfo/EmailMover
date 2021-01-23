Attribute VB_Name = "DirectInjection"

Sub InjectIntoDefectShooter()
EnvironementPath = CStr(Environ("USERPROFILE"))
StoragePath = "\\sourcesafe\CUSTO\_Custo\Projects\CBA\01-PDModels\02-Phase 2\040-Monitoring\040-Defects\Client-Details\R3"

Dim objItem As Object
Dim oInspector As Inspector
Dim defect As DefectDetails
Set DefectCollection = New Collection

'add the current window as well
    Set oInspector = Application.ActiveInspector
    If Not oInspector Is Nothing Then
        Set objItem = oInspector.CurrentItem
        AddSingleItem objItem, Nothing, DefectCollection, 0
        
        'use the key to find the correct email
        Set defect = DefectCollection.Item(1)
        
        'if a genuine defect then send to defect shooter
        If Not defect Is Nothing Then
             If SendToDefectShooter(defect) = True Then
                MsgBox "Injected", vbOKOnly, "Defect Shooter"
                
             Else
                 MsgBox "Not Injected", vbCritical + vbOKOnly, "Defect Shooter"
             End If
        End If
    End If
End Sub
