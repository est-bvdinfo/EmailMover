VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_colFolders As Collection
Private m_dicAlreadyAddedToStructure As Dictionary
Private m_dicFoundItem As Dictionary
Private m_jsonData As Dictionary
Private m_rootFolderTag As String
Private Sub Class_Terminate()

    Set m_colFolders = Nothing

End Sub
Public Property Get folders() As FolderList

    Set folders = m_colFolders

End Property
Public Property Get FolderTag() As String
    FolderTag = m_rootFolderTag
End Property

Public Sub CreateFolderStructure(dicJson As Object)
Dim folderInfo As FolderItem
Dim Item As Object
Dim key As Variant

Set m_dicFoundItem = New Dictionary
Set m_dicAlreadyAddedToStructure = New Dictionary
Set m_colFolders = New Collection

If TypeOf dicJson Is Dictionary Then
    Set m_jsonData = dicJson
    'first convert all the item in a class object then store them into a dictionary
    For Each Item In dicJson.Item("data")
   
        Set folderInfo = New FolderItem
        
            'build the item
            folderInfo.CreateFolder Item
            If (Not m_dicFoundItem.Exists(folderInfo.Id)) Then
                    m_dicFoundItem.Add folderInfo.Id, folderInfo
            End If
    Next

    ' populate the structure
    For Each key In m_dicFoundItem.keys
        AddToStructure m_dicFoundItem(key)

    Next
    
    
End If
Set Item = Nothing
Set dicJson = Nothing
End Sub
Public Sub PopulateTreeView(ByRef treeobject As clsTreeView)

'clean the tree
treeobject.NodesClear

'populate it
AddChildrenToTreeView treeobject, Nothing, m_colFolders

'display it
treeobject.Refresh

End Sub

Private Sub AddToStructure(Item As FolderItem, Optional parentItem As FolderItem)
Dim Id As Variant
Dim subFolderInfo As FolderItem
    If Not m_dicAlreadyAddedToStructure.Exists(Item.Id) Then
        If Not Item.ChildIds Is Nothing Then
            If Item.ChildIds.Count > 0 Then
                For Each Id In Item.ChildIds
                    ' find the children
                    Set subFolderInfo = m_dicFoundItem.Item(Id)
                    AddToStructure subFolderInfo, Item
                Next
            End If
        End If
            'add to the map to avoid duplicates
            m_dicAlreadyAddedToStructure.Add Item.Id, Item
            
            'skip the Recycle bin and unnecessary
            If Not ((Item.Scope = "WsRoot" And Item.ChildIds.Count = 0) Or (Item.Scope = "RbRoot")) Then
            
                If parentItem Is Nothing Then
                    m_rootFolderTag = Item.Id
                    m_colFolders.Add Item
                Else
                    parentItem.ChildFolders.Add Item
                End If
                
            End If
    End If

End Sub

Private Sub AddChildrenToTreeView(ByVal treeobject As clsTreeView, _
    ByVal parentNode As clsNode, ByVal folders As _
    Collection)

Dim Item As FolderItem
Dim new_node As clsNode

    ' Examine each XML child.
     For Each Item In folders
        ' Add the child to the TreeView.
        If parentNode Is Nothing And Not treeobject Is Nothing Then
            Set new_node = treeobject.AddRoot(Item.Id, Item.title)
        Else
            Set new_node = parentNode.AddChild(Item.Id, Item.title)
        End If
            new_node.Expanded = True
        
        'check if there is available child
         If Not Item.ChildFolders Is Nothing Then
                If Item.ChildFolders.Count > 0 Then
                    ' Add the child's children.
                    AddChildrenToTreeView Nothing, new_node, Item.ChildFolders
                End If
        End If

    Next Item
End Sub
