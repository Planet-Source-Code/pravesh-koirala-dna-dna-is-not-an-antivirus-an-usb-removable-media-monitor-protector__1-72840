VERSION 5.00
Begin VB.Form frmImmunize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Immunize"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisAutorun 
      Caption         =   "Disable Autorun"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Immunity"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdImmunize 
      Caption         =   "Immunize"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Device to Immunize"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
      Begin VB.OptionButton Options 
         Caption         =   "Drive with this Letter"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton Options 
         Caption         =   "My Local Drives"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Options 
         Caption         =   "USB Drives"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmImmunize.frx":0000
         Left            =   3000
         List            =   "frmImmunize.frx":0002
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmImmunize.frx":0004
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmImmunize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Dim RemDrive(10) As String
Dim LocalDrive(15) As String
Public IsActive As Boolean
Dim I As Integer

Private Sub cmdCancel_Click()
    'Remove immunization
    If MsgBox("You have chosed to remove immunization from a drive." & vbCrLf & _
          "It is not recommended. Do you still want to continue?", vbYesNo _
          Or vbExclamation) = vbYes Then
          
    'If only a single drive is selected..
        If Combo1.Enabled Then
            modImmunization.CancelImmunization (Combo1.List(Combo1.ListIndex))
            WT "Removed immunization of " & Combo1.List(Combo1.ListIndex)
 
    'If entire removable drive is selected...
        ElseIf Options(0).Value = True Then
    
            WT "Removed immunization of Removable drives"
            For I = 0 To UBound(RemDrive) - 1
                If RemDrive(I) <> "" Then modImmunization.CancelImmunization RemDrive(I)
            Next
        
    'If Local Drives are selected
        Else
            WT "Removed immunization of Local drives"
            For I = 0 To UBound(LocalDrive) - 1
                If LocalDrive(I) <> "" Then modImmunization.CancelImmunization LocalDrive(I)
            Next
        End If  'Combo1.enabled
        
    End If  'Msgbox
    MsgBox "DeImmunization successful!", vbOKOnly Or vbInformation
    CheckAvaivality
End Sub

Private Sub cmdDisAutorun_Click()
    Form1.m_dis_Click
End Sub

Private Sub cmdImmunize_Click()
    'Immunize a single drive with a fixed drive letter
    If Combo1.Enabled Then
        modImmunization.Immunize (Combo1.List(Combo1.ListIndex))
        WT "Immunized " & Combo1.List(Combo1.ListIndex)
    ElseIf Options(0).Value = True Then
        WT "Immunized Removal Drives"
        For I = 0 To UBound(RemDrive) - 1
            If RemDrive(I) <> "" Then modImmunization.Immunize RemDrive(I)
        Next
    Else
        WT "Immunized Local Drive"
        For I = 0 To UBound(LocalDrive) - 1
            If LocalDrive(I) <> "" Then modImmunization.Immunize LocalDrive(I)
        Next
    End If
    MsgBox "Immunization successful! Now you are safe", vbOKOnly Or _
            vbInformation, "Success"
    CheckAvaivality
End Sub

Private Sub Form_Load()
    Me.Icon = Form1.Icon
    WT "\\Now Entering Immunization Form", True
    Form1.m_exit.Enabled = False
    Form1.mnuImmunize.Enabled = False
    IsActive = True
    CheckAvaivality
End Sub

Private Sub CheckAvaivality()
    Dim I As Integer, Dl As String
    Dim RD As Integer, LD As Integer
    Combo1.Clear
    For I = 0 To 25
        Dl = Dl & Chr(Asc("A") + I) & ":\"
      Select Case GetDriveType(Dl)
        Case 3
            Combo1.AddItem Dl & IIf(modImmunization.CheckForImmunity(Dl) = True, _
                         " -Immunized!", "")
            LocalDrive(LD) = Dl
            LD = LD + 1
        Case 2
            Combo1.AddItem Dl & IIf(modImmunization.CheckForImmunity(Dl) = True, _
                         " -Immunized!", "")
            RemDrive(RD) = Dl
            RD = RD + 1
      End Select
      
        Dl = ""
    Next
    
    Combo1.ListIndex = 0
    WT "Updated the immunization list. There are a total of " & LD _
    & " local drives. And " & RD & " Removable drives."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.m_exit.Enabled = True
    Form1.mnuImmunize.Enabled = True
    IsActive = False
    WT "// Now Exiting from Immunization Form", False, True
End Sub

Private Sub Options_Click(Index As Integer)
    If Options(2).Value = True Then Combo1.Enabled = True Else _
                                    Combo1.Enabled = False
End Sub

Public Sub AddNewDrive(Dl As String)
    WT "Added Drive " & Dl & " for immunizing option"
    CheckAvaivality
End Sub
