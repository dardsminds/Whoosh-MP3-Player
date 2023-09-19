Attribute VB_Name = "modLVDrag"
Public Sub LVDragDropMulti(ByRef lvList As ListView, ByVal x As Single, ByVal y As Single)

    Dim objDrag As ListItem
    Dim objDrop As ListItem
    Dim objNew As ListItem
    Dim objSub As ListSubItem
    Dim intIndex As Integer
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim intSelected As Integer
    Dim arrItems() As ListItem
    
    On Error GoTo ErrHandlerDragDropMulti
    
    'Retrieve the original items
    Set objDrop = lvList.HitTest(x, y)
    Set objDrag = lvList.SelectedItem
    If (objDrop Is Nothing) Or (objDrag Is Nothing) Then
        Set lvList.DropHighlight = Nothing
        Set objDrop = Nothing
        Set objDrag = Nothing
        Exit Sub
    End If
    
    'Retrieve the drop position
    intIndex = objDrop.Index
    intCount = lvList.ListItems.Count
    intSelected = 0
    'Remove the drop highlighting
    Set lvList.DropHighlight = Nothing

    'Loop through and retrieve the selected items
    For intLoop = 1 To intCount
        If lvList.ListItems(intLoop).Selected Then
            intSelected = intSelected + 1
            ReDim Preserve arrItems(1 To intSelected) As ListItem
            Set arrItems(intSelected) = lvList.ListItems(intLoop)
        End If
    Next
    'Loop through in reverse and remove the selected items
    'Going in reverse prevents index shifting
    For intLoop = UBound(arrItems) To LBound(arrItems) Step -1
        lvList.ListItems.Remove arrItems(intLoop).Index
    Next
    'Loop through again and add the items back
    'Going in reverse keeps the items in order
    For intLoop = UBound(arrItems) To LBound(arrItems) Step -1
        Set objDrag = arrItems(intLoop)
        'Add it back into the dropped position
        Set objNew = lvList.ListItems.Add(intIndex, objDrag.Key, objDrag.text, objDrag.Icon, objDrag.SmallIcon)
        'Copy the original subitems to the new item
        If objDrag.ListSubItems.Count > 0 Then
            For Each objSub In objDrag.ListSubItems
                objNew.ListSubItems.Add objSub.Index, objSub.Key, objSub.text, objSub.ReportIcon, objSub.ToolTipText
            Next
        End If
        objNew.Selected = True
    Next
    
    'Destroy all objects
    ReDim arrItems(1)
    Set arrItems(1) = Nothing
    Set objNew = Nothing
    Set objDrag = Nothing
    Set objDrop = Nothing
    Exit Sub
ErrHandlerDragDropMulti:
    ErrorMsg Err.Number, Err.Description, Err.Source, True
End Sub

Public Sub LVDragDropSingle(ByRef lvList As ListView, ByVal x As Single, ByVal y As Single)

    Dim objDrag As ListItem
    Dim objDrop As ListItem
    Dim objNew As ListItem
    Dim objSub As ListSubItem
    Dim intIndex As Integer
    
    'Retrieve the original items
    Set objDrop = lvList.HitTest(x, y)
    Set objDrag = lvList.SelectedItem
    If (objDrop Is Nothing) Or (objDrag Is Nothing) Then
        Set lvList.DropHighlight = Nothing
        Set objDrop = Nothing
        Set objDrag = Nothing
        Exit Sub
    End If
    
    'Retrieve the drop position
    intIndex = objDrop.Index
    
    'Remove the dragged item
    lvList.ListItems.Remove objDrag.Index
    'Add it back into the dropped position
    Set objNew = lvList.ListItems.Add(intIndex, objDrag.Key, objDrag.text, objDrag.Icon, objDrag.SmallIcon)
    'Copy the original subitems to the new item
    If objDrag.ListSubItems.Count > 0 Then
        For Each objSub In objDrag.ListSubItems
            objNew.ListSubItems.Add objSub.Index, objSub.Key, objSub.text, objSub.ReportIcon, objSub.ToolTipText
        Next
    End If
    'Reselect the item
    objNew.Selected = True
    
    'Destroy all objects
    Set objNew = Nothing
    Set objDrag = Nothing
    Set objDrop = Nothing
    Set lvList.DropHighlight = Nothing

End Sub
