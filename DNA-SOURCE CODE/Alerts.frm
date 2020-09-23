VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alert!!! Drive may be infected."
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6690
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6690
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   4440
      Top             =   960
   End
   Begin VB.CommandButton cmdExtras 
      Appearance      =   0  'Flat
      Caption         =   "Extra Analysis>>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   840
      Width           =   1455
   End
   Begin VB.Frame frmActions 
      Caption         =   "Available Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
         Height          =   375
         Left            =   5280
         TabIndex        =   28
         Tag             =   "Deletes the exe file"
         ToolTipText     =   "Deletes the infected files"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Rename"
         Height          =   375
         Left            =   3560
         TabIndex        =   27
         ToolTipText     =   "Renames the files so that they may not be executed"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear Attrs."
         Height          =   375
         Left            =   1840
         TabIndex        =   26
         Tag             =   "Clicking this button will clear the attributes of the exe File"
         ToolTipText     =   "Sets the attributes of infected files as normal. This will make them visible."
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open Safe"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Opens the drive safely so that you may not get infected. First set the attributes to normal to view infected files"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdAv 
      Caption         =   "Pass 2 AV"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton command9 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   6240
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   "Probability of being a Virus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   6495
      Begin VB.Label Label9 
         Caption         =   "Suggested Action: "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label8 
         Caption         =   "Inf Killer's Rating (Out of 100): "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "About Binary File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   6495
      Begin VB.CheckBox Check 
         Caption         =   "Hidden"
         Height          =   255
         Index           =   7
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check 
         Caption         =   "Read Only"
         Height          =   255
         Index           =   6
         Left            =   3600
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check 
         Caption         =   "System"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox Check 
         Caption         =   "Archive"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "File Size "
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Name of File: "
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "About Autorun.inf"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6375
      Begin VB.TextBox Text1 
         Height          =   1095
         Left            =   2880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Text            =   "Alerts.frx":0000
         Top             =   150
         Width           =   3375
      End
      Begin VB.CheckBox Check 
         Caption         =   "Archive"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check 
         Caption         =   "System"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox Check 
         Caption         =   "Read Only"
         Height          =   255
         Index           =   1
         Left            =   240
         MaskColor       =   &H8000000F&
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check 
         Caption         =   "Hidden"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      Caption         =   $"Alerts.frx":0006
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label2 
      Caption         =   "Drive Label:  "
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Type:  "
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Drive Letter:  "
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "File System:  "
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   3975
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DNA by Devil Labs. Start with the readme file.
'Many things in this code are from Internet. We thank all whose articles helped
'We write this whole code and ask for their forgiveness for We didn't mentioned their
'name here.

Option Explicit
''''''''''''''''''''''''''''''''''''APIS'''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
        "GetVolumeInformationA" (ByVal lpRootPathName As String, _
        ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
        ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, _
        ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
        ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias _
        "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias _
        "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
        lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString _
        As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias _
        "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, _
        ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Const DRIVE_CDROM As Long = 5
Private Const DRIVE_FIXED As Long = 3
Private Const DRIVE_REMOVABLE As Long = 2
Private Const SND_ASYNC As Long = &H1
Private Const SND_MEMORY As Long = &H4
Private Const SND_NODEFAULT = &H2
Private Const Flagsd& = SND_ASYNC Or SND_NODEFAULT

'*******Variables************************
Dim SST As String           'SST is declared for rough use
Dim Bin_Files(10) As String
Public Action As Boolean, ShowMessage As Boolean
Public IsActive As Boolean, DT As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Private Sub Check_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If anyone clicks the checkboxes that show the attributes of the files
    'there will be no effect of it.
    Check(Index).Value = IIf(Check(Index).Value = 1, 0, 1)
End Sub

Private Sub CmdAv_Click()
    'Passing the file to the Antivirus whose path is given
    Dim path_Av As String
    path_Av = GetSt("AVPath", "")
    
    If path_Av <> "" Then
        path_Av = path_Av & " " & Bin_Files(0)
        Close #12
        Shell path_Av, vbNormalFocus
    Else
        MSG "Path of Antivirus is not provided. Go to options and give path.", _
        vbInformation Or vbOKOnly
    End If
    
End Sub

Private Sub cmdDetail_Click()
    
    If cmdDetail.Caption = "<<" Then    'If previously under high detail
        SaveSt "Detail", "Low"
        cmdDetail.Caption = ">>"
        Unload Me
        Me.Show
    Else
        SaveSt "Detail", "High"
        cmdDetail.Caption = "<<"
        Unload Me
        Me.Show
    End If
    
End Sub

Private Sub cmdExtras_Click()
    FrmExtraAnalysis.Show , Me
End Sub

Private Sub cmdExtras_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
    'Delete all Files
    WT "User chose to Delete the Files ."
On Error GoTo a:
    SetAttr Bin_Files(0), vbNormal
    SetAttr File, vbNormal
    Close #12
    Kill Bin_Files(0)
    Kill File
    Action = True
    MSG "Performed requested Action", vbOKOnly Or vbCritical
    SetChecks
Exit Sub

a:
    SST = err.Description
    WT "Can't Delete the file. " & SST
    MSG "Can't set file attribute to normal." & SST & vbCrLf & _
    "Maybe it is a CD/DVD ROM.", vbInformation Or vbOKOnly
    
End Sub


Private Sub Command2_Click()
On Error GoTo a:
    Command7_Click
    WT "User chose to rename the file"
    Close #12
    Name File As File & ".DNA"
    Name Bin_Files(0) As Bin_Files(0) & ".DNA"
    File = File & ".DNA"
    Bin_Files(0) = Bin_Files(0) & ".DNA"
    Action = True
Exit Sub

a:
    SST = err.Description
    WT "Can't Delete the file. " & SST
    MSG "Can't set file attribute to normal." & SST & vbCrLf & _
    "Maybe it is a CD/DVD ROM.", vbInformation Or vbOKOnly
    
End Sub

Private Sub Command3_Click()
    Shell "notepad.exe " & File, vbMaximizedFocus
End Sub

Private Sub Command5_Click()
    Shell "explorer.exe " & Dl, vbMaximizedFocus
End Sub

Private Sub Command7_Click()
On Error GoTo a:
    WT "User chose to set attribute of Files to normal"
    SetAttr Bin_Files(0), vbNormal
    SetAttr File, vbNormal
    MSG "Performed requested Action", vbOKOnly Or vbInformation
    SetChecks
Exit Sub

a:
    SST = err.Description
    WT "Can't Delete the file. " & SST
    MSG "Can't set file attribute to normal." & SST & vbCrLf & _
    "Maybe it is a CD/DVD ROM.", vbInformation Or vbOKOnly
    
End Sub

Private Sub Command8_Click()
    
    WT "//Dialog closed. User closed the dialog.", False, True
    Unload Me
End Sub

Private Sub Command9_Click()

    WT "//Dialog closed. User stopped the checking.", False, True
    i_thread = 2
    command9.Visible = False
    Unload Me
End Sub

Private Sub Form_Activate()
    i_thread = 0
    IsActive = True
End Sub

Private Sub Form_Load()
On Error GoTo errs:
    
    
    
    'If the setting is to automatically fix the drive then
    'We disable error and information messages And try to fix the drive
    'Else we continue
    
    If GetSt("Automatic", "no") = "yes" And isCheckAll = False Then
        ShowMessage = False
        If DT = 5 Then GoTo exitS:
        If GetSt("Delete") = "True" Then Command1_Click
        If GetSt("Rename") = "True" Then Command2_Click
        modImmunization.Immunize Dl
exitS:
        Unload Me
Exit Sub
    Else
        ShowMessage = True
    End If
    
    DT = GetDriveType(Dl)

    Me.Icon = Form1.Icon
    
    WT "Entered Dialog"
    
    Action = False          'This action variable determines if user did smt with drive
    Form1.m_exit.Enabled = False
    Form1.m_Each.Enabled = False
    
    command9.Visible = isCheckAll
    
    'Set window as topmost
    SetWindowPos Me.hWnd, -1, (Screen.Width - Me.Width) \ 30, (Screen.Height - Me.Height) \ 30, Me.ScaleWidth, Me.ScaleHeight, 1
    
    'Get volume information
    Dim ame As String, tem As String, sz As Long, a As Long
    ame = Space$(255)
    tem = Space$(255)
    Label3.Caption = Label3.Caption & UCase(Dl)
    GetVolumeInformation Dl, ame, 20, 0, a, a, tem, 20
    Label2.Caption = Label2.Caption & ame
    Label4.Caption = Label4.Caption & tem
    Rem_Char tem

    'Load Autorun.inf contents into the TextBox
    LoadFileInTextBox
    
    'Executable file
    FindAndStoreExeFiles

    WT "Now checking the drive for probability of danger"
    If Bin_Files(0) = "" Then Label5.Caption = Label5.Caption & "  FILE MISSING!!!!!!"
    SST = Bin_Files(0)
    Label5.Caption = Label5.Caption & Bin_Files(0)
    Bin_Files(0) = SST

    Label6.Caption = Label6.Caption & (FileLen(Bin_Files(0)) \ 1024) & " Kb"
    SetChecks
    WT "File name is " & SST & " And size is " & FileLen(Bin_Files(0)) & "."
    'End the running executable

    'calculate possibility
    Dim pro As Integer
    pro = pro + IIf(Check(0).Value = 1, 10, 0)
    pro = pro + IIf(Check(1).Value = 1 And DT <> 5, 10, 0)
    pro = pro + IIf(Check(2).Value = 1, 10, 0)
    pro = pro + IIf(pro = 30, 10, 0)
    pro = pro + IIf(Check(5).Value = 1, 10, 0)
    pro = pro + IIf(Check(6).Value = 1 And DT <> 5, 10, 0)
    pro = pro + IIf(Check(7).Value = 1, 10, 0)
    pro = pro + IIf(Check(0).Value = 1 And Check(7).Value = 1, 10, 0)
    pro = pro + IIf(Check(2).Value = 1 And Check(5).Value = 1, 10, 0)
    pro = pro + IIf(Check(1).Value = 1 And Check(6).Value = 1 And DT <> 5, 10, 0)
    GetPrivateProfileString "autorun", "shell\open\command", "", tem, 255, File
    Rem_Char tem
    pro = pro + IIf(tem = "", 0, 30)
    GetPrivateProfileString "autorun", "shell\explore\command", "", tem, 255, File
    Rem_Char tem
    If Not tem = "" Then tem = FindActualPAth(tem, Dl)
    pro = pro + IIf(tem = "", 0, 30)
    pro = pro + IIf(pro = 150, 40, 0)
    pro = pro \ 2
    WT "Probability checked: It is " & pro
    
    
    Select Case pro
        Case Is = 100
            tem = str(pro) & " (Critical!!!)"
            ame = "Delete both the autorun and executable!!!!"
        Case Is > 70
            tem = str(pro) & " (Very Dangerous)"
            ame = "Rename the files OR Quarantine them."
        Case Is > 30
            tem = str(pro) & " (Dangerous)"
            ame = "Pass them to Antivirus if installed or qurantine them"
        Case Is > 0
            tem = str(pro) & " (Less Dangerous)"
            ame = "Set attributes of the Executable and the inf file to normal " _
                    & "and inspect urself"
        Case Else
            tem = str(pro) & " (Shouldn't be Dangerous)"
            ame = "If it is from a trusted source then it isn't dangerous."
    End Select
    
    Label8.Caption = Label8.Caption & tem
    Label9.Caption = Label9.Caption & ame

    'Set Drive Type
    Select Case DT
        Case 5
             tem = "CD\DVD ROM"
             'With Cd roms we can't edit or delete infected files
            Command7.Enabled = False
            Command2.Enabled = False
            Command1.Enabled = False
        Case 3
            tem = "Fixed Disk (Hard drive)"
        Case 2
            tem = "Removable Disk (USB, Mem card)"
        Case Else
            tem = "Can't Determine the drive type"
    End Select
    
    Label1.Caption = Label1.Caption & tem
    checkForDetails
    sndPlaySound App.path & "\file\sound.wav", Flagsd&

On Error GoTo err2:
    'IF the current account is of administrator then we can even lock
    'a usb drive or Cd-rom but not a hard disk
    'Drive or file is locked because if user accesses the infected file
    'while the alert dialog is on, he may get infected
    Sleep 500
    If Administrator And (DT = 2 Or DT = 5) Then
        Open "//./F:" For Binary Access Read Write Lock Read Write As #12
        'Open "//./" & Left(Dl, 2) For Input Lock Read Write As #12
        WT "Locked the drive " & Dl
    Else
        Open Bin_Files(0) For Input Lock Read Write As #12
        WT "Locked the File " & Bin_Files(0)
    End If

    If ModPictureAndFolders.ListUpdateRootOnly(Dl) Then Timer1.Enabled = True
    
    

Exit Sub
    
    'Err2 is only accesed via the locking mechanism. Drive needs some time to
    'finish what it is doing and hence drive may not be locked. So we shall
    'Resume the same statement until the drive finishes its job and isn't locked
err2:
    If err.Number = 70 Then Resume
    
errs:
    
    WT "Error occured in the Dialog: " & err.Description
    Resume Next
End Sub
Private Sub SetChecks()
On Error Resume Next
    Dim Xs As VbFileAttribute
    Xs = GetAttr(Bin_Files(0))
    Check(4).Value = IIf(Xs And vbArchive, 1, 0)
    Check(5).Value = IIf(Xs And vbSystem, 1, 0)
    Check(6).Value = IIf(Xs And vbReadOnly, 1, 0)
    Check(7).Value = IIf(Xs And vbHidden, 1, 0)
    Xs = GetAttr(File)
    Check(0).Value = IIf(Xs And vbHidden, 1, 0)
    Check(1).Value = IIf(Xs And vbReadOnly, 1, 0)
    Check(2).Value = IIf(Xs And vbSystem, 1, 0)
    Check(3).Value = IIf(Xs And vbArchive, 1, 0)
End Sub

Private Sub Rem_Char(ByRef Stri As String)
On Error Resume Next
    Stri = Left(Stri, InStr(1, Stri, vbNullChar, vbBinaryCompare) - 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
        
    If i_thread = 0 Then i_thread = 1
    Form1.m_exit.Enabled = True
    Form1.m_Each.Enabled = True

    IsActive = False
    If Action Then MsgBox "To apply the changes to the drive you need to unplug and " _
                    & "replug the device. If changes are made to a hard disk," & _
                    vbCrLf & "then restart. Do it now", vbInformation Or vbOKOnly
    Extra = Dl
    Dl = ""
    Close #12
End Sub

Private Sub MSG(MS As String, Optional Sty As VbMsgBoxStyle)
    If ShowMessage = False Then Exit Sub
    MsgBox MS, Sty Or vbSystemModal, "DNA"
End Sub



Private Sub checkForDetails()
'Change the appearance of the dialog according to high or low details.
    Dim Tp As Integer
    Tp = 2550
    
    If GetSt("Detail", "Low") = "Low" Then
        Dim var As Variant
        
            For Each var In Me.Controls
                If TypeOf var Is Frame Then var.Visible = False
            Next
            
        Me.Height = 3500
        frmActions.Visible = True
        frmActions.Top = Tp - 900
        command9.Top = Tp
        Command8.Top = Tp
        CmdAv.Top = Tp
        cmdDetail.Top = Tp
        cmdDetail.Caption = ">>"
    End If
    
End Sub


Private Sub Timer1_Timer()
    'Timer is used to blink the Extra Analysis button so that user may get attracted
    cmdExtras.Visible = Not cmdExtras.Visible
End Sub

Private Sub FindAndStoreExeFiles()
    Dim Counter As Integer
        
        For Counter = 0 To 10
            Bin_Files(Counter) = Space(50)
        Next
        
    GetPrivateProfileString "autorun", "open", "", Bin_Files(0), 50, File
    Rem_Char Bin_Files(0)
    Bin_Files(0) = FindActualPAth(Bin_Files(0), Dl)
    
    GetPrivateProfileString "autorun", "shell\open\command", "", Bin_Files(1), 50, File
    Bin_Files(1) = FindActualPAth(Bin_Files(1), Dl)
    If Bin_Files(0) = "" Then Bin_Files(0) = Bin_Files(1)
    
    GetPrivateProfileString "autorun", "shell\explore\command", "", Bin_Files(2), 50, File
    Bin_Files(2) = FindActualPAth(Bin_Files(2), Dl)
    If Bin_Files(0) = "" Then Bin_Files(0) = Bin_Files(2)
    
End Sub

Private Function FindActualPAth(ByVal str As String, ByVal path As String) As String
    Dim Length As Integer, I As Integer, tmp As String
    I = 1
    str = path & str
    
    If existS(str) Then
        FindActualPAth = str
Exit Function
    End If

Redo:
    Length = InStr(I, str, " ")
    If Length <> 0 Then tmp = Left(str, Length - 1) Else Exit Function
    
    If existS(tmp) Then
        If Not (GetAttr(tmp) And vbDirectory) Then
            FindActualPAth = tmp
        Else
            I = Length + 1
            GoTo Redo:
        End If
    Else
        I = Length + 1
        GoTo Redo:
    End If
    
End Function

Private Sub LoadFileInTextBox()
'Formerly richtextbox was used but since it needed an ocx which may not be
'with everyone and including it was just unnecessary because it had no other
'use than just displaying the Autorun entries
    Dim Length As Long, I As Long, fName As String
    Dim bt() As Byte, tmp As String
    
    fName = File
    Length = FileLen(fName)
    ReDim bt(Length - 1)
    Open fName For Binary As #1
    Get #1, 1, bt()
    
    For I = 0 To Length - 1
        tmp = tmp & Chr(bt(I))
    Next
    
    Close #1
    Text1.text = tmp
End Sub
