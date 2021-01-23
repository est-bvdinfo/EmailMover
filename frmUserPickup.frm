VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserPickup 
   Caption         =   "Wrike"
   ClientHeight    =   8175
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   16740
   OleObjectBlob   =   "frmUserPickup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUserPickup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'##########Treeview Code##########
'Add this to your form's declaration section
Private WithEvents mcTree As clsTreeView
Attribute mcTree.VB_VarHelpID = -1
Private mbExit As Boolean    ' to exit a SpinButton event
'/##########Treeview Code##########


Public AppName As String

#If Mac Then
    Const mcPtPixel As Long = 1
#Else
    Const mcPtPixel As Single = 0.75
#End If

Private Sub CmdPushToWrike_Click()
Dim i As Integer
Dim key As String
Dim oKey As Variant
Dim defect As DefectDetails
Dim task As WrikeTask
Dim tTitle As String
Dim tDescription As String
Dim jsonTask As String


Dim selectedEmails As Collection
Dim selectedFolders As Collection

Set selectedEmails = New Collection
Set selectedFolders = New Collection

LblStatus.Caption = "Collect selection"

'get the selected email
For i = 0 To LstMails.ListCount - 1
    If (LstMails.Selected(i) = True) Then
        key = LstMails.List(i, 0)
        selectedEmails.Add key
    End If
Next i

' get selected folders
For i = 0 To LstSelect.ListCount - 1
    key = LstSelect.List(i, 1)
    selectedFolders.Add key
Next i

'fetch the defects properties from the initial datacollection
For Each oKey In selectedEmails
    'use the key to find the correct email
    Set defect = DefectCollection.Item(oKey)
    
    'build wrike object
    If Not defect Is Nothing Then
        LblStatus.Caption = "Process: " & defect.DefectId
        
        Set task = New WrikeTask
        tTitle = defect.DefectId + "# " + defect.title
        
        'build wrike tasks
        task.Initialize tTitle, defect.Description, selectedFolders
        
        LblStatus.Caption = "Send to wrike: " & defect.DefectId
        
        'inject into wrike
        If SendTaskToWrike(task) = True Then
            'if succeeded then remove from the list
             LblStatus.Caption = "Removing mail: " & defect.DefectId
                Dim mail As Variant
                For Each mail In selectedEmails
                    For i = 0 To LstMails.ListCount - 1
                        If (LstMails.List(i, 0) = mail) Then
                            LstMails.RemoveItem i
                            Exit For
                        End If
                    Next i
                Next
        Else
            MsgBox "Not able to inject defect" & defect.DefectId & vbCrLf & "Please check logs in temp folder", vbExclamation, "Task to Wrike Fail"
        End If
    End If
Next
LblStatus.Caption = ""


End Sub

Private Sub CmdSendToDefectShooter_Click()
Dim i As Integer
Dim key As String
Dim oKey As Variant
Dim defect As DefectDetails
Dim tTitle As String
Dim tDescription As String



Dim selectedEmails As Collection
Dim selectedFolders As Collection

Set selectedEmails = New Collection
Set selectedFolders = New Collection

LblStatus.Caption = "Collect selection"

'get the selected email
For i = 0 To LstMails.ListCount - 1
    If (LstMails.Selected(i) = True) Then
        key = LstMails.List(i, 0)
        selectedEmails.Add key
    End If
Next i



'fetch the defects properties from the initial datacollection
For Each oKey In selectedEmails
    'use the key to find the correct email
    Set defect = DefectCollection.Item(oKey)
    
    'build wrike object
    If Not defect Is Nothing Then
        LblStatus.Caption = "Process: " & defect.DefectId
        
        tTitle = defect.DefectId + "# " + defect.title
        
        LblStatus.Caption = "Send to DefectShooter: " & defect.DefectId
        
        'inject into wrike
        If SendToDefectShooter(defect) = True Then
            'if succeeded then remove from the list
             LblStatus.Caption = "Removing mail: " & defect.DefectId
                Dim mail As Variant
                For Each mail In selectedEmails
                    For i = 0 To LstMails.ListCount - 1
                        If (LstMails.List(i, 0) = mail) Then
                            LstMails.RemoveItem i
                            Exit For
                        End If
                    Next i
                Next
              'display to the user which item has been sent
              LstSelect.AddItem defect.DefectId & "(" & defect.ExternalID & ")" & ":" & defect.title
                
        Else
            MsgBox "Not able to inject defect" & defect.DefectId & vbCrLf & "Please check logs in temp folder", vbExclamation, "Injection failed"
        End If
    End If
Next
LblStatus.Caption = ""

'copy to sourcesafe

CopyAllFiles EnvironementPath & "\Documents\temp\", "\\sourcesafe\CUSTO\_Custo\Projects\CBA\01-PDModels\02-Phase 2\040-Monitoring\040-Defects\Client-Details\R3"

End Sub


Private Sub CmdRemove_Click()
Dim selectedEmails As Collection
Set selectedEmails = New Collection
Dim i As Integer
Dim key As Variant
'get the selected email
For i = 0 To LstMails.ListCount - 1
    If (LstMails.Selected(i) = True) Then
        key = LstMails.List(i, 0)
        selectedEmails.Add key
    End If
Next i

'remove from list
Dim mail As Variant
For Each mail In selectedEmails
    For i = 0 To LstMails.ListCount - 1
        If (LstMails.List(i, 0) = mail) Then
            LstMails.RemoveItem i
            Exit For
        End If
    Next i
Next

End Sub

Private Sub UserForm_Activate()
Set mcTree = New clsTreeView


EnvironementPath = CStr(Environ("USERPROFILE"))
StoragePath = "\\sourcesafe\CUSTO\_Custo\Projects\CBA\01-PDModels\02-Phase 2\040-Monitoring\040-Defects\Client-Details\R3"

''id the tree
'mcTree.AppName = "Wrike"
'mcTree.CheckBoxes(bTriState:=False) = True

'RetrieveFoldersAndLoadTree mcTree

With Me.LstSelect
    .ColumnCount = 2
    .ColumnWidths = "1000;0"
End With

'load the selected email list
LoadSelectedEmails LstMails

With Me.LstMails
    .ColumnCount = 5
    .ColumnWidths = "0;50;50;1000;200"
End With





End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub


Private Sub mcTree_NodeCheck(cNode As clsNode)
Dim currentPos As Integer
If cNode.Checked Then
    LstSelect.AddItem

    If (LstSelect.ListCount > 0) Then
        currentPos = LstSelect.ListCount - 1
    End If
    
    LstSelect.List(currentPos, 1) = cNode.key
    LstSelect.List(currentPos, 0) = cNode.FullPath

Else
    Dim i As Integer
    For i = LstSelect.ListCount - 1 To 0 Step -1
          If (LstSelect.List(i, 1) = cNode.key) Then
            LstSelect.RemoveItem (i)
          End If
    Next i

End If


End Sub

Private Sub LoadSelectedEmails(ByRef listOfmails As ListBox)

Dim objItem As Object
Dim i As Integer

Set DefectCollection = New Collection
 
 For Each objItem In ActiveExplorer.Selection
    AddSingleItem objItem, listOfmails, DefectCollection, i
 Next objItem



End Sub



