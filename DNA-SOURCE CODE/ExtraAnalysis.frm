VERSION 5.00
Begin VB.Form FrmExtraAnalysis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   5280
   End
   Begin VB.ListBox List1 
      Height          =   3210
      ItemData        =   "ExtraAnalysis.frx":0000
      Left            =   120
      List            =   "ExtraAnalysis.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   2040
      Width           =   6135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Fix"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Compare"
      Height          =   375
      Left            =   5190
      TabIndex        =   6
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select All"
      Height          =   375
      Left            =   4020
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6600
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2850
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   $"ExtraAnalysis.frx":0004
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Icon"
      Height          =   315
      Left            =   6720
      TabIndex        =   4
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "  Files"
      Height          =   195
      Left            =   6600
      TabIndex        =   1
      Top             =   3480
      Width           =   405
   End
End
Attribute VB_Name = "FrmExtraAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Stop_Update As Boolean
Dim MyDl As String
Private Sub Command1_Click()
    'Start Scanning
    Command4.Enabled = True
    Command1.Enabled = False
    Stop_Update = False
    
    List1.Clear
    List2.Clear
    
    Label4.Caption = "Now Scanning Drive Please Wait....."
    
    ListUpdate MyDl
    
    FilTerFiles
    
    If List1.ListCount = 0 Then
        Stop_Update = True
        Unload Me
        Exit Sub
    End If
    
    Label1.Caption = List1.ListCount & " Files"
    WT Label1.Caption & " Suspected"
    List1_Click         'Show the icon of the first file
    Command4.Enabled = False
    Command1.Enabled = True
    Label4.Caption = ""
    Stop_Update = True
End Sub

Private Sub FilTerFiles()
    'Filter files having a matching folder from list2 to list1
On Error Resume Next

    Dim st() As String, Length As Long
    ReDim st(List2.ListCount - 1)
    
    For I = 0 To List2.ListCount - 1
        st(I) = List2.List(I)
    Next
    
    List2.Clear
    
    For I = 0 To UBound(st)
        Length = 0
        Length = FileLen(st(I))
        
        If Length Then
            List1.AddItem (st(I))
        End If
    Next
    
End Sub

Private Sub Command3_Click()
    Dim status As Boolean

    If Command3.Caption = "Select All" Then
        status = True
        Command3.Caption = "DeSelect All"
    Else
        status = False
        Command3.Caption = "Select All"
    End If
    
    For I = 0 To List1.ListCount - 1
        List1.Selected(I) = status
    Next
    
End Sub


Private Sub Command4_Click()
    'Stop the process of updating
    Stop_Update = True
End Sub

Private Sub Command5_Click()
'Compare the Selected files with other files on the basis of their sizes
    Dim Y As String, Length As Long, X As Long
    Dim FirstAttr As FileAttribute
    X = -1
    
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) = True Then
            FirstAttr = GetAttr(Left(List1.List(I), Len(List1.List(I)) - 4))
            X = I
    Exit For
        End If
    Next
    
If X = -1 Then Exit Sub

X = 0           'Now X will be used to count the selected files
Length = FileLen(List1.List(X))

    For I = 0 To List1.ListCount - 1
        Y = List1.List(I)
        
        If FileLen(Y) = Length And GetAttr(Left(Y, Len(Y) - 4)) = FirstAttr Then
            List1.Selected(I) = True
            X = X + 1
        End If
        
    Next
            
    WT "Compared files. " & X & " files selected according to first selection."
End Sub

Private Sub Command6_Click()
    'Fix all Checked Files
    Dim Y As String
    Dim I As Integer
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then
            Y = List1.List(I)
            SetAttr Left(Y, Len(Y) - 4), vbNormal
            Kill Y
        End If
    Next
    
    MsgBox "All files fixed. You are safe.", vbOKOnly Or vbInformation, "DNA"
    List1.Clear
    WT "Fixed " & I & " files"
    
End Sub

Private Sub Form_Load()
    Me.Icon = Form1.Icon
    MyDl = Dl
    WT "\\Scanning for executable folders of " & MyDl, True
    Me.Caption = "Analysis of drive " & MyDl
    Timer1.Enabled = True
    
    If GetDriveType(MyDl) = 5 Then Command6.Enabled = False
    
    Form1.m_exit.Enabled = False
    Form1.mnuExtra.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If Stop_Update = False Then
        Cancel = 1
        Exit Sub
    End If
    
    i_thread = 1
    WT "//Analysis completed", False, True
    Form1.m_exit.Enabled = True
    Form1.mnuExtra.Enabled = True
End Sub

Private Sub List1_Click()
    RetrieveIcon List1.List(List1.ListIndex), Picture1
End Sub
 
Private Sub ListUpdate(ByVal path As String)
    'All entries are first added in a hidden listbox list2 and filtered
    'in second step and added in list1
    Dim Buf() As String
    
    On Error GoTo err:
    
    RetrieveAllFolders path, Buf
    
    For I = 0 To UBound(Buf)
        If Stop_Update Then Exit Sub
        DoEvents
        List2.AddItem Buf(I) & ".exe"
        ListUpdate Buf(I) & "\"
    Next
    
err:
End Sub

Private Sub Timer1_Timer()
    'Timer is used to hold the execution of check so that the form may be properly displayed
    Timer1.Enabled = False
    Command1_Click
End Sub


