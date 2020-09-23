Attribute VB_Name = "modImmunization"
'Here we immunize the removable drive (Not CDS).
'We create a non replacable autorun.inf file. So that the virus may not
'create the entry.
'THe Autorun.inf folder contains DNA_Dir.. folder which can't be replaced without
'extra tweaks.
'This module is called primarily by frmImmunize

Option Explicit
Dim I As Integer
Public ImmunizedFolderName As String

Public Function Immunize(Dl As String) As Boolean
    Dim X As Long, StrToPut As String
    Dl = Left(Dl, 3)
    If CheckForImmunity(Dl) Then Exit Function
On Error GoTo errOccured:
    'If there is not an autorun file already then error occurs and we go to
    'erroccured position where we create a new autorun directory
    Dim FN As String
    FN = Dl & "Autorun.inf"
    X = FileLen(FN)
    SetAttr FN, vbNormal
    Kill FN

errOccured:
    'Remove previous error handler and create a new one
    On Error GoTo ex:
    MkDir FN
    'Create Readme file
    Open FN & "\DNA_Readme.txt" For Output As #1
    Print #1, "This Immunized folder has been created By DNA 1.0.0" & vbCrLf _
           & "This can only be removed by same program." & vbCrLf _
           & "CAUTION!!! Don't Remove SIG.DNA File!!!!!!!"

    Close #1
    Open FN & "\SIG.DNA" For Binary As #1
    Dim Data(100) As Byte, Seq As String
    Seq = CreateRandomSequence
    StrToPut = "DON't Delete This File!!!. It is Extremely necessary for DNA. " & _
           "Sig_Start " & Seq & " "
    
    For I = 0 To Len(StrToPut) - 1
        Data(I) = Asc(Mid(StrToPut, I + 1, 1))
    Next
    
    Put #1, 1, Data
    Close #1
    SetAttr FN & "\sig.dna", vbHidden Or vbSystem Or vbReadOnly
    MkDir FN & "\DNA_Dir...\"
    SetAttr FN, vbHidden Or vbReadOnly Or vbSystem
    Immunize = True
    
Exit Function

ex:
    'we will get here if file doesn't exists
    StrToPut = err.Description
    MsgBox "Err occured " & StrToPut, vbCritical Or vbOKOnly, "DNA"
    WT "Err occured " & StrToPut
End Function

Public Function CancelImmunization(Dl As String) As Boolean

    Dim X As Long, MsgRes As String, FN As String
    'Deleting the Autorun directory.
On Error GoTo errOccured:
    Dl = Left(Dl, 3)
    X = FileLen(Dl & "autorun.inf\SIG.DNA")
    If X <> 101 Then MsgRes = MsgBox("The SIG file has been modified. DNA maynot be able to DE-Immunize it" & vbCrLf _
                        & "Do you still want to give a try? Success isn't guranteed" _
                        , vbExclamation Or vbYesNo) Else GoTo DOIT
    If MsgRes = vbYes Then GoTo DOIT

Exit Function

errOccured:

Exit Function

DOIT:
    modDeviceInfo.RetrieveSequence Dl
    'Remove Sequence entry from the Drive.
    
    
    FN = Dl & "autorun.inf"
    RmDir FN & "\DNA_Dir...\"
    Kill FN & "\DNA_readme.txt"
    SetAttr FN & "\Sig.dna", vbNormal
    Kill FN & "\sig.dna"
    SetAttr FN, vbNormal
    RmDir Dl & "\autorun.inf"
    CancelImmunization = True
End Function

Public Function CheckForImmunity(DLetter As String) As Boolean
    On Error Resume Next
    Dim X As Long
    X = FileLen(DLetter & "autorun.inf\sig.dna")
    CheckForImmunity = X
End Function

