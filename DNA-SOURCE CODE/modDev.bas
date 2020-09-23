Attribute VB_Name = "modDev"
'Retrieve the device information etc from LParam passed from main module
'Also contains extra functions whose functionality can be guessed by function's name.
'WT means write Text (To Logfile)
Option Explicit

Public Type DEV_BROADCAST_HDR
dbch_size As Long
dbch_devicetype As Long
dbch_reserved As Long
End Type

Public Type DEV_BROADCAST_VOLUME
dbch_size As Long
dbch_devicetype As Long
dbch_reserved As Long
dbch_unitmask As Long
dbch_flags As Integer
End Type

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                        (ByRef Destination As Any, ByRef Source As Any, _
                        ByVal Length As Long)
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias _
                        "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As _
                        String, ByRef lpFreeBytesAvailableToCaller _
                        As PUlarge_Integer, ByRef lpTotalNumberOfBytes As _
                        PUlarge_Integer, ByVal lpTotalNumberOfFreeBytes As Long) _
                        As Long

Private Type PUlarge_Integer
Low As Long
High As Long
End Type

Public Administrator As Boolean, IsInForm As Boolean

Public Function FindChanges(lp As Long) As Integer

    'Find the Drive letter from lp.
    'MSDN can be consulted to know abt dev_brodcast_[volume][hdr]
    Dim X As DEV_BROADCAST_VOLUME
    Dim Y As DEV_BROADCAST_HDR
    Dim Res As Integer
    
    CopyMemory Y, ByVal (lp), Len(Y)
    
    If Y.dbch_devicetype = 2 Then       'DBT_DEVTYP_VOLUME
        CopyMemory X, ByVal (lp), Len(X)
        
        'dbch_unitmask contains Drive letter number in power of two.
        '    driveletter = 2^x
        'or, log(driveletter) = log (2^x)
        'or, log(driveletter) = x log 2
        'or, x = log(driveletter)/ log 2
        
        Res = Log(X.dbch_unitmask) / Log(2)
        FindChanges = Res + 65
    End If
    
End Function

Public Sub WT(st As String, Optional EnteringForm As Boolean = False _
                            , Optional LeavingForm As Boolean = False)
On Error Resume Next
    If LeavingForm Then IsInForm = False
    If IsInForm Then st = "  " & st
    Print #14, st
    If EnteringForm Then IsInForm = True
End Sub

Public Sub SaveSt(Key As String, Data As String)
    SaveSetting "DNA", "Settings", Key, Data
    WT "Saved a setting. Key name = " & Key & ". Data given = " & Data
End Sub

Public Function GetSt(Key As String, Optional Default As String = "") As String
    GetSt = GetSetting("DNA", "Settings", Key, Default)
    WT "Retrieved a setting. Key name = " & Key & ". Data = " & GetSt
End Function

Public Function GetUSedSpace(ByVal Dl As String) As Double

    Dim Available_Bytes As PUlarge_Integer, TotalBytes As PUlarge_Integer
    Dim Y As Double
    
    GetDiskFreeSpaceEx Dl, Available_Bytes, TotalBytes, ByVal 0
    
    Y = GEtDbl(TotalBytes.Low, TotalBytes.High)
    
    If Y = 0 Then       'If drive is empty then
        GetUSedSpace = -1
        Exit Function
    End If
    
    Y = Y - GEtDbl(Available_Bytes.Low, Available_Bytes.High)
    Y = Y / 1073741824      '1073741824 = 1024 * 1024 * 1024
    
    GetUSedSpace = Y
End Function

Private Function GEtDbl(low_part As Long, high_part As Long) As Double
    
    Dim result As Double
    
    result = high_part
    
    If high_part < 0 Then result = result + 2 ^ 32
    
    result = result * 2 ^ 32
    result = result + low_part
    
    If low_part < 0 Then result = result + 2 ^ 32
    
    GEtDbl = result
End Function

Public Function IsAdministrator() As Boolean
'Only Administrators are allowed to use drives as files.
    
    IsAdministrator = IsUserAdmin
    
    WT "Checked if current account is of the administrator. " & IIf(IsAdministrator, _
                    "Yes I am.", "No I'm not.")
                    
End Function

'This is to Protect DNA's files from being modified by any infector virus
Public Sub File_Protection()
'This is reserved for Future work
End Sub


Public Function existS(st As String)

On Error GoTo error:
    FileLen (st)
    existS = True
error:
    
End Function


