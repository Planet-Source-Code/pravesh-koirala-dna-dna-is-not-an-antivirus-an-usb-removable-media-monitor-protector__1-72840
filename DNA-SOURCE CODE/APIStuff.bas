Attribute VB_Name = "APIStuff"
Option Explicit
Public Declare Function RegisterWindowMessage Lib "user32.dll" Alias _
                    "RegisterWindowMessageA" (ByVal lpString As String) As Long
                    
Public Declare Function GetDriveType Lib "kernel32.dll" Alias _
                    "GetDriveTypeA" (ByVal nDrive As String) As Long
                    
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" _
                    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
                    
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, _
                    ByVal hWndInsertAfter As Long, ByVal X As Long, _
                    ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
                    ByVal wFlags As Long) As Long
                    
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                    (ByVal hWnd As Long, ByVal lpOperation As String, _
                    ByVal lpFile As String, ByVal lpParameters As String, _
                    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                            (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
                            ByVal MSG As Long, ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                            (ByVal hWnd As Long, ByVal nIndex As Long, _
                            ByVal dwNewLong As Long) As Long
Public Declare Function IsUserAdmin Lib "SetupApi.dll" () As Boolean

Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_LBUTTONDBLCLK As Long = &H203

Private Const WM_CONTEXTMENU = &H7B
Private Const WM_PASTE = &H302
Private Const DBT_DEVICEARRIVAL As Long = 32768         '&H8000
Private Const WM_DEVICECHANGE As Long = &H219
Private Const dbt_devtype_volume As Long = &H2
Private Const DBT_DEVICEQUERYREMOVE As Long = 32769     '&H8001
Private Const DBT_DEVICEREMOVECOMPLETE As Long = 32772  '&H8004

Public status As Boolean        'Whether the monitor is active or not
Public OldWindowProc As Long    'Stores the oldwindow procedure of the app
Public i_thread As Long         'USed to see if a dialog is active or not.
Public isCheckAll As Boolean
Public Extra As String
Public Action_Extra As Boolean
Public Dl As String             'The drive letter of the inserted drive


Public File As String           'Consists the full path of the Autorun
Dim a As Integer                'USed to find the drive letter
Public uMsg As Long             'Variable for receiving QueryCancelAutoplay message

Public Function NewWindowProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo a:
    If MSG = WM_MYHOOK Then
        Select Case lParam
         Case &H205 'Right button up in Tray icon
            Form1.POP
           Case NIN_BALLOONUSERCLICK    'Balloon clicked
                Immunize
        End Select
    End If
    
    If MSG = uMsg Then              'Received QueryCancelAutorun
        WT "Received window message and suppressed autorun."
        NewWindowProc = 1
    End If
    
    'A new Drive Added or removed
    If MSG = WM_DEVICECHANGE And lParam <> 0 And status = True Then
    Debug.Print Hex(wParam)
     If wParam = DBT_DEVICEARRIVAL Then
        a = FindChanges(lParam)
        
        If a > 50 And a < 100 Then
            Dl = Chr(a)
            WT "Found a new Device. Drive Letter is " & Dl
            GetIfPrese True
        End If
        
     ElseIf wParam = DBT_DEVICEQUERYREMOVE And Dl <> "" Then
        a = FindChanges(lParam)
        
        If Chr(a) = Left(Dl, 1) Then
            NewWindowProc = 0
            Exit Function
        End If
     
     ElseIf wParam = DBT_DEVICEREMOVECOMPLETE And Dl <> "" Then
        a = FindChanges(lParam)
        If Chr(a) = Left(Dl, 1) Then
            Unload Dialog
        End If
     
     End If
   End If
    NewWindowProc = CallWindowProc( _
            OldWindowProc, hWnd, MSG, wParam, _
            lParam)
            Exit Function
a:
   WT "Error occured in SubClassing Sub" & err.Description
End Function

Public Sub GetIfPrese(Optional showballon As Boolean)
'if showballon = true, drive has been added recently (not startupScan or similar)
  On Error GoTo errs:
  Dim Length As Long
  
  Dl = Dl & ":\"
  
  If frmImmunize.IsActive Then frmImmunize.AddNewDrive (Dl)
  If showballon Then WT "Now checking Drive " & Dl & "for infection"
  
  File = Dl & "autorun.inf"
  On Error GoTo a:      'If file doesn't exists error will occur
  Length = FileLen(File)
  On Error GoTo errs: 'Now reinstall the previous err handler
  
  If modImmunization.CheckForImmunity(Dl) Then
    'This drive is immunized. Check for other things only
    DoExtraAnalysis Not (showballon)
    Exit Sub
  End If
  
  'Drive not immunized and contains Autorun entry.
  'Show the alert dialog.
  i_thread = 0
  WT "\\Drive " & Dl & " Suspected. Now entering Dialog", True
  Dialog.Show , Form1
  Exit Sub
  'Autorun doesn't exists. Check for executable folders and prompt to immunize
a:
  On Error GoTo errs:
  DoExtraAnalysis Not (showballon)
  'If the drive isn't CdROM and showballoon is true then prompt with  Balloon
  If showballon And GetDriveType(Dl) <> 5 Then
    TrayModify ("New Removal Drive Detected." & vbCrLf & _
                 "Click here to immunize it.")
    WT "Drive wasn't infected. Displayed tip for immunization."
  End If
  
  Exit Sub
errs:
  WT "Error occured in the GetIfPreseSub" & err.Description
End Sub

Public Sub Immunize()
  If Dl = "" Then Dl = Extra
  modImmunization.Immunize (Dl)
End Sub

Public Sub DoExtraAnalysis(IsStartupScan As Boolean)
'For checking executable folders, the whole drive should be checked.
'Costs a huge time if drive is a hard disk or has a lot of space.
'Only the drives with less space should be checked (By default 5 GB but can be
'modified in options).
  Dim Size As Double
  On Error GoTo err:
If IsStartupScan Then

    Size = modDev.GetUSedSpace(Dl)
    If Size = -1 Then Exit Sub              'Size = -1 = Drive is empty :-)
    'IF drive shows some hint of being infected then listupdaterootonly becomes
    'True
    If Not (ModPictureAndFolders.ListUpdateRootOnly(Dl)) Then Exit Sub
    If Size > Val(GetSt("ExtraSize", "5")) Then
        If MsgBox("During scan, drive " & Dl & " showed some suspicious behaviour." & _
                "It should be further checked thoroughly." & vbCrLf & _
                "But it is greater than the allowed size. Checking may take a long time" & _
                vbCrLf & "Do you wish to thoroughly scan the drive?", vbYesNo Or vbInformation, "Alert") _
                = vbYes Then GoTo DoScan Else Exit Sub
    End If
    GoTo DoScan
Else
    GoTo DoScan
End If
Exit Sub

DoScan:
Action_Extra = True
i_thread = 0
FrmExtraAnalysis.Show , Form1
Action_Extra = False
Exit Sub

err:
WT "Error occured while doing extra analysis " & err.Description
End Sub
