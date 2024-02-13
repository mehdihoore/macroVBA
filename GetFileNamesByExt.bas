Function GetFileNamesByExt(FolderPath As String, FileExt As String) As Variant
    Dim fso As Object, fld As Object, fil As Object
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(FolderPath)

    ReDim arr(1 To fld.Files.Count) As Variant
    i = 1

    For Each fil In fld.Files
        If LCase(Right(fil.Name, Len(FileExt))) = LCase(FileExt) Then
            arr(i) = fil.Name
            i = i + 1
        End If
    Next fil

    GetFileNamesByExt = arr
End Function
