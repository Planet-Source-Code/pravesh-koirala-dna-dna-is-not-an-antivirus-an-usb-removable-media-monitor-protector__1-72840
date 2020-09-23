VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DNA Settings"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "The Antivirus Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   700
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Enter the path for the AV that is installed in your computer. It will be used to scan suspected files."
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   5880
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "When infected files are found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   5295
      Begin VB.Frame Frame5 
         Caption         =   "Do What?"
         Height          =   855
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   3495
         Begin VB.OptionButton OptRename 
            Caption         =   "Rename infected files"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton optDelete 
            Caption         =   "Delete infected files"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.OptionButton OptWarning 
         Caption         =   "Show Warning Box"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   4575
      End
      Begin VB.OptionButton optAtm 
         Caption         =   "Automatically Fix them  (Only USB Devices not CDs)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   4935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Choose one of the following"
         Enabled         =   0   'False
         Height          =   855
         Left            =   600
         TabIndex        =   4
         Top             =   1680
         Width           =   3495
         Begin VB.OptionButton optLD 
            Caption         =   "Dialog box with less details"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   2535
         End
         Begin VB.OptionButton Opthd 
            Caption         =   "Dialog box with greater details"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Value           =   -1  'True
            Width           =   2775
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Top             =   900
         Width           =   495
      End
      Begin VB.CheckBox chkAutorun 
         Caption         =   "Disable Autorun for all drives (Strictly Recommended)"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   4455
      End
      Begin VB.CheckBox chkStartup 
         Caption         =   "Start DNA when window starts"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "GB"
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Scan for executable folders of drives upto"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oshell As IWshShell_Class
Private Sub Check1_Click()

End Sub

Private Sub Command1_Click()
'Write the values of entities in registry
On Error Resume Next
    Dostartup
    DoAutorun
    
    If optAtm.Value = True Then
        SaveSt "Automatic", "yes"
        If optDelete.Value = True Then SaveSt "Delete", "True" Else _
                              SaveSt "Delete", "False"
        If OptRename.Value = True Then SaveSt "Rename", "True" Else _
                              SaveSt "Rename", "False"
    Else
        SaveSt "Automatic", "no"
        If optLD.Value = True Then SaveSt "Detail", "Low" Else _
                                SaveSt "Detail", "High"
    End If
    SaveSt "AVPath", Text1.text
    SaveSt "ExtraSize", Text2.text
    Unload Me
End Sub

Private Sub DoAutorun()
    If chkAutorun.Value = 1 Then
        oshell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveTypeAutoRun", 255, "REG_DWORD"
        oshell.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveTypeAutoRun", 255, "REG_DWORD"
    Else
        oshell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveTypeAutoRun"
        oshell.RegDelete "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer\NoDriveTypeAutoRun"
    End If
    
   'MsgBox "Requested Action was performed", vbOKOnly Or vbInformation
    WT "Disabled autorun."

End Sub

Private Sub Dostartup()
    If chkStartup.Value = 0 Then
        oshell.RegDelete "HKLM\Software\microsoft\windows\currentversion\run\InfKiller"
        WT "Removed Me from startup"
    Else
        oshell.RegWrite "HKLM\Software\microsoft\windows\currentversion\run\InfKiller", App.path & "\infkiller.exe", "REG_SZ"
        WT "Added Me to startup"
    End If
    'MsgBox "Requested Action was performed", vbOKOnly Or vbInformation
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    'Show dialog and take the path of antivirus
    cd1.Filter = "Executables(*.exe)|*.exe"
    cd1.ShowOpen
    Text1.text = IIf(cd1.FileName = "", Text1.text, cd1.FileName)
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Icon = Form1.Icon
    WT "\\Now entering Options dialog", True
    Form1.m_exit.Enabled = False
    Form1.mnuopt.Enabled = False
    Set oshell = New IWshShell_Class
    chkStartup.Value = IIf(oshell.RegRead("HKLM\Software\microsoft\windows\" _
                        & "currentversion\run\InfKiller") = "", 0, 1)
    chkAutorun.Value = IIf(oshell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\" _
                       & "CurrentVersion\policies\Explorer\NoDriveTypeAutoRun") = 255 _
                        , 1, 0)

    optAtm.Value = IIf(GetSt("Automatic") = "yes", True, False)
    OptWarning.Value = Not optAtm.Value
    optLD.Value = IIf(GetSt("Detail") = "Low", True, False)
    Opthd.Value = Not optLD.Value
    optDelete.Value = IIf(GetSt("Delete", "False") = "True", True, False)
    OptRename.Value = IIf(GetSt("Rename", "False") = "True", True, False)
    Text1.text = GetSt("AVPath")
    Text2.text = GetSt("ExtraSize", "5")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oshell = Nothing
    Form1.m_exit.Enabled = True
    Form1.mnuopt.Enabled = True

    WT "//Leaving the options form", False, True
End Sub

Private Sub optAtm_Click()
    Frame3.Enabled = False
    Frame5.Enabled = True
    optLD.Enabled = False
    Opthd.Enabled = False
    optDelete.Enabled = True
    OptRename.Enabled = True
End Sub


Private Sub OptWarning_Click()
    Frame3.Enabled = True
    Frame5.Enabled = False
    optLD.Enabled = True
    Opthd.Enabled = True
    optDelete.Enabled = False
    OptRename.Enabled = False
End Sub

Private Sub Text2_Validate(Cancel As Boolean)

    If Len(Trim(str(Val(Text2.text)))) <> Len(Text2.text) Then
        MsgBox "Text not allowed. Re-Enter the text in box", vbCritical
        Cancel = True
    End If
    
End Sub
