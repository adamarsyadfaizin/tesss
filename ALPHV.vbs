Dim fso, folder, file, subFolder
Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = "D:\\NoddieJar"

If fso.FolderExists(folderPath) Then
    On Error Resume Next ' Lewati error jika ada file sedang digunakan
    
    Set folder = fso.GetFolder(folderPath)
    
    ' Hapus semua file
    For Each file In folder.Files
        fso.DeleteFile file.Path, True
    Next
    
    ' Hapus semua subfolder
    For Each subFolder In folder.SubFolders
        fso.DeleteFolder subFolder.Path, True
    Next
    
    ' Bersihkan objek
    Set folder = Nothing
End If

Set fso = Nothing
