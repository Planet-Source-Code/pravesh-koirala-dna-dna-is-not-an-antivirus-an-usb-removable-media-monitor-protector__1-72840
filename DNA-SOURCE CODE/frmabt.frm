VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About DNA"
   ClientHeight    =   4410
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6015
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3043.861
   ScaleMode       =   0  'User
   ScaleWidth      =   5648.396
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6015
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Devil's DNA 1.0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage your Usb and other removable drives easily."
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4560
      TabIndex        =   0
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      BorderWidth     =   2
      X1              =   0
      X2              =   4394.762
      Y1              =   2567.611
      Y2              =   2567.611
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "For extra Details read the Readme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Credits for development of this software goes to Devil Labs. Thanks to other respective owners for their multimedia stuffs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "A quality product from devil labs. Install it and it will monitor your removable media and prevents your pc from being infected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
    Form1.m_exit.Enabled = True
    Form1.m_a.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = Form1.Icon
    WT "\\User wants to know about me. About form Displayed", True
    Form1.m_exit.Enabled = False
    Form1.m_a.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WT "//Leaving the form", False, True
End Sub

