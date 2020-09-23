Attribute VB_Name = "ModPictureAndFolders"
'This module is used for retrieving a file's icon and putting into a
'picturebox. APIs can be consulted in MSDN Library or Internet.
'Also it is used to get all files and sub.folders of a folder
'The RetrieveAllFolders is used by the FrmExtraAnalysis form. It calls the
'Function recursively and collects all folders only.
Option Explicit
Private Const SHGFI_SYSICONINDEX = &H4000, SHGFI_SMALLICON = &H1, ILD_TRANSPARENT = &H1
Private Type SHFILEINFO
       hIcon As Long
       iIcon As Long
       dwAttributes As Long
       szDisplayName As String * 255
       szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
      (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, _
      ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "Comctl32.dll" _
(ByVal himl As Long, ByVal I As Long, ByVal hDCDest As Long, ByVal X As Long _
, ByVal Y As Long, ByVal Flags As Long) As Long

Private shinfo As SHFILEINFO, sshinfo As SHFILEINFO
Dim I As Long
Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime1 As Long: ftCreationTime2 As Long
    ftLastAccessTime1 As Long: ftLastAccessTime2 As Long
    ftLastWriteTime1 As Long: ftLastWriteTime2 As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long

Public Sub RetrieveIcon(fName As String, DC As PictureBox)
    'Retrievinf icon of fname and putting in DC picturebox
    If fName = "" Then Exit Sub
    Dim hImgLarge As Long   'the handle to the system image list
    DC.Cls
    hImgLarge = SHGetFileInfo(fName, 0&, shinfo, Len(shinfo), _
                              SHGFI_SYSICONINDEX)
    Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC.hDC, 0, 0, ILD_TRANSPARENT)
End Sub

Public Sub RetrieveAllFolders(ByVal path As String, Buffer() As String, Optional _
            ByVal fileAttribute As Long = vbDirectory)
    'Returns all folders of path and stores in Buffer variable.
    'It could have been accomplished by a folder list box but it doesn't provides
    'folders which are hidden

    Dim temp(255) As String, WinFData As WIN32_FIND_DATA
    Dim J As Long
On Error GoTo ERR:
    Dim SearchHandle As Long, I As Integer, temps As String
    'Create a handle and retrieve the first file
    SearchHandle = FindFirstFile(path & "*.*", WinFData)
    
    'IS the file a directory?
    If WinFData.dwFileAttributes And fileAttribute Then
    
        temps = Left$(WinFData.cFileName, InStr(1, WinFData.cFileName, vbNullChar _
                               ) - 1)
                               
        'You must debug this part to see what the below code does
        If (temps <> "." And temps <> "..") Then
            temp(I) = path & temps
            I = I + 1
        End If
        
    End If
    
    While FindNextFile(SearchHandle, WinFData) = 1
        
        'Check if the file is a directory
        If WinFData.dwFileAttributes And fileAttribute Then
                temps = Left$(WinFData.cFileName, InStr(1, WinFData.cFileName, vbNullChar _
                              ) - 1)
                              
                If (temps <> "." And temps <> "..") Then
                    temp(I) = path & temps
                    I = I + 1
                End If
                
        End If
        
    Wend
    
    'Close the handle
    FindClose SearchHandle
    
    Dim ActualData() As String
    ReDim ActualData(I - 1)
    
    For J = 0 To I - 1
        ActualData(J) = temp(J)
    Next
    
    Buffer = ActualData
ERR:
End Sub

Public Function ListUpdateRootOnly(ByVal Dl As String) As Boolean
    'Check only the root folder of the drive for existence of executable folders
    'This provides a hint of whether the drive contains executable folders
    Dim Buf() As String, Length As Long
    
On Error Resume Next

    RetrieveAllFolders Dl, Buf
    
    For I = 0 To UBound(Buf)
        Length = FileLen(Buf(I) & ".exe")
        
        If Length Then
            ListUpdateRootOnly = True
Exit Function
        End If
        
    Next
    
End Function


