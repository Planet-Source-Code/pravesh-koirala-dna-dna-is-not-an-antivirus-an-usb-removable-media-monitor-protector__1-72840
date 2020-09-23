Attribute VB_Name = "modDeviceInfo"
'Functionality of this module is reserved for Future Use
'It shall be used on further releases. Most probably the next Release.
'What we do here is use the 15-character sequence to store datas about the
'drive in hard disk. So that it may be used some other time.
'#If 0 Then

Option Explicit
Public Type SequenceData
 Label As String * 20
 Size As Long
 LastChecked As Date
 Sequence As String * 15
 DeviceType As Integer
End Type
Dim Seq_Data As SequenceData
Public Function RetreiveSequenceData(Seq As String) As SequenceData
    Dim Fl As String, Fl_size As Long, Fl_records As Integer
    Dim I As Integer
    
    Fl = "File.dat"
    Open Fl For Random As #1
    
    Fl_size = FileLen(Fl)
    Fl_records = Fl_size / Len(Seq_Data)
    
    For I = 1 To Fl_records
        Get #1, I, Seq_Data
        
        If Seq_Data.Sequence = Seq Then
            RetreiveSequenceData = Seq_Data
Exit Function
        End If
        
    Next
    
Close #1
End Function

Public Function AddSequenceData(ByRef DataToAdd As SequenceData)
    Dim Fl As String, Fl_size As Long, Fl_records As Integer
    
    Fl = "File.dat"
    
    Open Fl For Random As #1
    Fl_size = FileLen(Fl)
    Fl_records = Fl_size / Len(DataToAdd)
    
    Put #1, Fl_records + 1, DataToAdd
    Close #1
End Function

Public Function RetrieveSequence(Dl As String) As String

On Error Resume Next

    Dim Sigfile As String, FN As Long
    Dim Data(14) As Byte
    Dim X As String
    Dim I As Integer
    
    Sigfile = Dl & "Autorun.inf\sig.dna"
    FN = FreeFile
    
    Open Sigfile For Binary As FN
    Get FN, 73, Data
    
    For I = 0 To 14
        X = X & Chr(Data(I))
    Next
    
    Close FN
    
    RetrieveSequence = X
End Function

Public Function CreateRandomSequence() As String
    'Here we create a random sequence. This sequence can be used to identify a
    'Drive. This functionality is currently not added but will be in future.
    Dim Seq As String, FileSeq As String
    Dim X As Integer, I As Integer

Redo:
    Randomize Timer
    
    For I = 1 To 15
        X = Fix((Rnd * 15) + 0)
        Seq = Seq & Hex(X)
    Next
    
    If Not existS(App.path & "\File\QuarantineData") Then _
            MkDir App.path & "\File\QuarantineData"
    
    On Error GoTo err:
    
    Open App.path & "\File\QuarantineData\SeQuenceList.dna" For Input As #5
    
    'Now checking if the created sequence is already in use. Probability = infinitely low
    
    While Not EOF(5)
        Input #5, FileSeq
        If Seq = FileSeq Then GoTo Redo:
    Wend
    Close #5

err:
    CreateRandomSequence = Seq
    
    Close #5
    
    Open App.path & "\File\QuarantineData\SeQuenceList.dna" For Append As #5
    Print #5, Seq
    Close #5
    
End Function

'#End If
