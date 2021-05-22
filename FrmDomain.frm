VERSION 5.00
Begin VB.Form FrmDomain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Domain Selection"
   ClientHeight    =   1995
   ClientLeft      =   5445
   ClientTop       =   3705
   ClientWidth     =   5115
   Icon            =   "FrmDomain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox List1 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label LblDomain 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Domain Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "FrmDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
    StrDomain = List1
    Unload Me
    FrmNetMessage.Show
End Sub
Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Dim NameSpace As IADsContainer
    Dim Domain As IADs
    Set NameSpace = GetObject("WinNT:")
    For Each Domain In NameSpace
        List1.AddItem Domain.Name
    Next
    List1.ListIndex = 0
End Sub
