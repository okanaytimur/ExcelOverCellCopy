Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Dim DataObject As Object
    Set DataObject = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObject.SetText Selection.Value
    DataObject.PutInClipboard

End Sub