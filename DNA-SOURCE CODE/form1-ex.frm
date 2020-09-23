VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6855
   Icon            =   "form1-ex.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "This form is the main form. It always remains hidden."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Menu m_opp 
      Caption         =   "Options"
      Tag             =   "Options|(Checked=0)(Enabled=-1)(Visible=0)(WindowList=0)"
      Begin VB.Menu m_tog 
         Caption         =   "Enabled"
         Tag             =   "Enabled|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuNothing 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtra 
         Caption         =   "Check A Drive For Exe Folders"
      End
      Begin VB.Menu m_Each 
         Caption         =   "Check Each Drive For Virus"
         Tag             =   "Check Each Drive For Virus|#se|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnuNothing2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImmunize 
         Caption         =   "Immunization"
      End
      Begin VB.Menu mnuopt 
         Caption         =   "Options..."
      End
      Begin VB.Menu m_dis 
         Caption         =   "Disable Autorun"
         Tag             =   "Disable Autorun|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu m_startup 
         Caption         =   "Execute on startup"
         Tag             =   "Execute on startup|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Visible         =   0   'False
      End
      Begin VB.Menu aaa 
         Caption         =   "-"
         Tag             =   "-|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu m_a 
         Caption         =   "About"
         Tag             =   "About|@""""Settings|#at|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu m_hlp 
         Caption         =   "Help"
         Tag             =   "Help|@""""Settings|#hp|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu assda 
         Caption         =   "-"
         Tag             =   "-|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu m_exit 
         Caption         =   "Exit"
         Tag             =   "Exit|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Be careFull While debugging. Improper debugging can cause a crash due to subclassing
'Never stop execution of program when it is being debugged.

Public oshell As IWshShell_Class    'For Reading Registry
Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private Sub Form_Load()
    
    If App.PrevInstance Then
        MsgBox "This app is already running!! You can find it in tray.", _
            vbOKOnly Or vbInformation
        WT "App is already running. Application is being end"
        End
    End If

On Error GoTo errs:
    
    
    Dim ascd As Integer                     ' A variable to store temporary values
    Set oshell = New IWshShell_Class

    'Opening log file
    ChDir App.path
    
    If existS("APP_log.log") Then
        If FileLen("APP_Log.log") / 1024 > 5000 Then Kill "APP_log.log"
    End If
    
    Open "APP_Log.log" For Append As #14
    WT "App started on " & Now

    uMsg = RegisterWindowMessage("QueryCancelAutoPlay")
    
    'Now we remove the error handler because if the program is not added
    'in run subkey the we get a error
    
On Error Resume Next

    m_startup.Checked = Not IsNull(oshell.RegRead("HKLM\Software\microsoft\windows\currentversion\run\InfKiller"))
    WT "Am I in startup? " & IIf(m_startup.Checked, "Yes", "No")

    'Error Handler added again
On Error GoTo errs:

    
    'Checking Whether we are administrator
    Administrator = IsAdministrator

    'Checking if it is the first time program is running
    ascd = Val(GetSt("IsFirst", "0"))
    
    If ascd = 0 Then
        SaveSt "IsFirst", "1"
        ascd = oshell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveTypeAutoRun")
        ascd = ascd Or oshell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveAutoRun")
        ascd = ascd Or oshell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveTypeAutoRun")
        ascd = ascd Or oshell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveAutoRun")
      
        If ascd <> 255 Then
            ascd = MsgBox("Hello User! Since this program is running for the first " & _
                            "time, DNA suggests you edit Options." & vbCrLf & _
                            "Also you must disable the autorun. Do you want to do it now?.", _
                            vbYesNo Or vbInformation)
                            
            If ascd = vbYes Then Call m_dis_Click
            
        End If
        
    End If

    Me.Hide
    TrayAdd
    WT "Created application's Tray"
 
    'Enable the monitor
    m_tog_Click

    'Subclassing this application. If you don't know abt subclassing then do some googling.
    OldWindowProc = SetWindowLong(Me.hWnd, -4, AddressOf NewWindowProc)
    WT "********************Started subclassing successfully.***************************"
    m_Each_Click                                   'Check each drive during startup
    
Exit Sub

errs:
    WT "Error occured: " & err.Description
Resume Next

End Sub

Public Sub POP()
    PopupMenu m_opp
End Sub


Private Sub m_a_Click()
    frmAbout.Show , Me
End Sub

Public Sub m_dis_Click()
    Dim temp As String
    
On Error GoTo err:

    oshell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveAutoRun", 255, "REG_DWORD"
    oshell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveTypeAutoRun", 255, "REG_DWORD"
    oshell.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveAutoRun", 255, "REG_DWORD"
    oshell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveTypeAutoRun", 255, "REG_DWORD"
  
    MsgBox "Autorun has been disabled. But it will only show effect after a restart." & _
                vbOKOnly Or vbInformation
    WT "Disabled autorun"
    
Exit Sub

err:

Debug.Print err.Number

    If err.Number = 70 Then
        temp = "Error occured while disabling Autorun." & _
                "You must log in as Administrator"
    Else
        temp = "Error occured while disabling Autorun." & err.Description
    End If
    
    MsgBox temp, vbCritical Or vbOKOnly
    WT temp
    
End Sub

Private Sub m_Each_Click()
    WT "Now checking each Drive."
    isCheckAll = True
    i_thread = 1
    'Here i_thread variable is used to pause the execution of this loop so that other
    'dialogs can be displayed one by one. If more drives are infected then you don't
    'get two or more dialogs at once but one by one.
    Dim DriveType As Byte
    
    For I = 0 To 25
        
        Dl = Chr(Asc("A") + I)
        DriveType = GetDriveType(Dl & ":\")
        
        If Not (DriveType = 0 Or DriveType = 1) Then GetIfPrese
a:
        DoEvents
        
        If i_thread = 0 Then
            GoTo a:
        ElseIf i_thread = 2 Then
            Dl = ""
            isCheckAll = False
Exit Sub
        End If
    Next
    
    isCheckAll = False
    Dl = ""
    
End Sub

Private Sub m_exit_Click()
  If MsgBox("Do you really want to Exit?", vbInformation Or vbYesNo) = vbYes Then _
        Unload Me
End Sub

Private Sub m_hlp_Click()
    WT "Readme File opened"
    ShellExecute Me.hWnd, "open", "Readme.txt", "", App.path, vbMaximizedFocus
End Sub

Private Sub m_startup_Click()

    If m_startup.Checked = True Then
    
        If MsgBox("It is not recommended to do so. Do you still want to remove this program from startup?", vbYesNo Or vbExclamation) = vbYes Then
            oshell.RegDelete "HKLM\Software\microsoft\windows\currentversion\run\InfKiller"
            WT "Removed Me from startup"
        End If
        
    Else
    
        oshell.RegWrite "HKLM\Software\microsoft\windows\currentversion\run\InfKiller", App.path & "\infkiller.exe", "REG_SZ"
        WT "Added Me to startup"
        
    End If
    
    m_startup.Checked = Not (m_startup.Checked)
    MsgBox "Requested Action was performed", vbOKOnly Or vbInformation
    
End Sub

Private Sub m_tog_Click()

    status = Not status
    WT "Status changed. Currently I am " & IIf(status, "Enable", "Disable")
    m_tog.Checked = status
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

        
        TrayRemove
        WT "Shutting down the DNA on " & Now
        WT ""
        WT ""
        Close #14
End Sub


Private Sub mnuExtra_Click()
    Dim DriveLetter As String, DriveType As Byte
    
    DriveLetter = InputBox("Enter the Drive Letter of the drive to test. eg: A:\" _
                , "DNA", "")
        
    If DriveLetter = "" Then Exit Sub
    If Right(DriveLetter, 1) <> "\" Then DriveLetter = DriveLetter & "\"
    
    DriveType = GetDriveType(DriveLetter)
    If DriveType = 0 Or DriveType = 1 Then Exit Sub
    
    If Right(DriveLetter, 1) <> "\" Then DriveLetter = DriveLetter & "\"
    
    Dl = DriveLetter
    FrmExtraAnalysis.Show , Me
    
End Sub

Private Sub mnuImmunize_Click()
    frmImmunize.Show , Me
End Sub

Private Sub mnuopt_Click()
    Form2.Show , Me
End Sub
