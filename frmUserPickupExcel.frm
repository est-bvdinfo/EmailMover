VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUserPickupExcel 
   Caption         =   "Wrike"
   ClientHeight    =   8175
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   16740
   OleObjectBlob   =   "frmUserPickupExcel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUserPickupExcel"
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


Private Sub CmdSelectAll_Click()
On Error GoTo Err_Handler
    'Purpose:   Select all items in the multi-select list box.
    'Return:    True if successful
    'Author:    Allen Browne. http://allenbrowne.com  June, 2006.
    Dim lngRow As Long

    If LstMails.MultiSelect Then
        For lngRow = 0 To LstMails.ListCount - 1
            LstMails.Selected(lngRow) = True
        Next
    End If

Err_Handler:
    Exit Sub


End Sub

Private Sub CmdUploadComments_Click()
Dim i As Integer
Dim key As String
Dim oKey As Variant
Dim defect As DefectDetails
Dim task As WrikeTask
Dim tTitle As String
Dim tDescription As String
Dim jsonTask As String
Dim commentSheet As Worksheet
Dim selectedEmails As Collection
Dim selectedFolders As Collection
Dim position As Integer

Set commentSheet = Sheets("ListOfComments")
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
        
        'find if task exists in wrike
        If (FindTaskInWrike(task) = True) Then
            LblStatus.Caption = "Find comments of " & defect.DefectId
            'inject into wrike
                Dim comment As WrikeComment
                For Each comment In task.GetComments
                    ' print comment
                    position = position + 1
                    commentSheet.Cells(position, 1) = defect.DefectId
                    commentSheet.Cells(position, 2) = comment.AuthorName
                    commentSheet.Cells(position, 3) = comment.text
                    commentSheet.Cells(position, 4) = comment.UpdatedDate
                    commentSheet.Cells(position, 5) = defect.ExternalID
                 Next
            ' then remove from the list
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
            MsgBox "Can't find this task in wrike :" & defect.DefectId
        End If
    
    End If
Next
LblStatus.Caption = ""

End Sub

Private Sub UserForm_Activate()


Set mcTree = New clsTreeView

Set mcTree.TreeControl = Me.frTreeControl

'id the tree
'mcTree.AppName = "Wrike"
'mcTree.CheckBoxes(bTriState:=False) = True

'RetrieveFoldersAndLoadTree mcTree

With Me.LstSelect
    .ColumnCount = 2
    .ColumnWidths = "1000;0"
End With

'load the selected email list
LoadFilteredRows Me.LstMails

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

Private Sub LoadFilteredRows(ByRef listOfmails)
Dim defect As DefectDetails
Dim objItem As Range
Dim key As String
Dim i As Integer
Dim visibleRange As Range


Set DefectCollection = New Collection

Set visibleRange = Range("A2:A3000")

 
 For Each objItem In visibleRange.SpecialCells(xlCellTypeVisible)

        key = Trim(objItem.Columns("A").Value)
        
        If (Len(key) >= 2) Then
            
            ' instanciate the defect based on the email
            Set defect = New DefectDetails
            defect.InitialiseFromExcel objItem.EntireRow
            
            key = defect.DefectId & "_" & i
            DefectCollection.Add defect, key
            
            'add to listbox
            listOfmails.AddItem
            
            listOfmails.List(i, 0) = key
            listOfmails.List(i, 1) = defect.DefectId
            listOfmails.List(i, 4) = defect.Severity
            listOfmails.List(i, 2) = defect.ExternalID
            If (Len(defect.title) > 230) Then
                listOfmails.List(i, 3) = Left(defect.title, 230)
            Else
                listOfmails.List(i, 3) = defect.title
            End If
             'increment next
            i = i + 1
        End If

  Next objItem

Set defect = Nothing
End Sub


