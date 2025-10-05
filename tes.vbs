' popup_prioritas_fix.vbs
Do
    response = MsgBox("Your OS are depracted.", _
        vbYesNo + vbSystemModal + vbCritical, "WARNING")
Loop While response <> vbYes

MsgBox "ERROR 503.", vbInformation + vbSystemModal, "fatal : error"
