VERSION 5.00
Begin VB.Form FrmNetMessage 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Net Messenger"
   ClientHeight    =   2100
   ClientLeft      =   4560
   ClientTop       =   4140
   ClientWidth     =   5775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Interval        =   60
      Left            =   1800
      Top             =   1440
   End
   Begin VB.PictureBox AniGif 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5040
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox ImgTorch2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   4920
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   945
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Send to"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OptSingle 
         BackColor       =   &H00000000&
         Caption         =   "Single"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00000000&
         Caption         =   "All"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "Send Message"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type the username of recipient or select the computername from the list below"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "FrmNetMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal As Boolean
Dim SysInfo As New SysInfo
Private Declare Function NetMessageBufferSend Lib _
  "NETAPI32.DLL" (yServer As Any, yToName As Byte, _
  yFromName As Any, yMsg As Byte, ByVal lSize As Long) As Long
Private Const NERR_Success As Long = 0&
Dim IntFlame As Integer
Dim StrComputer, StrUser, strMsg As String

Private Sub SendAll()
    GetFullName
    For Each Computer In Container
        RetVal = SendMessage(Computer.Name, "Me", strMsg & vbNewLine & "Thank You," & vbNewLine & StrUser)
    Next
End Sub
Private Sub SendSingle()
    GetFullName
    RetVal = SendMessage(Combo1, "Me", strMsg & vbNewLine & "Thank You," & vbNewLine & StrUser)
End Sub
Private Sub CmdSend_Click()
    strMsg = InputBox("Type the Message Text you wish to send", "Net Messenger")
    If OptAll.Value = True Then
        Dim Response As Long
        Response = MsgBox("Are you sure you want to send " & Chr(34) & strMsg & Chr(34) & " to Every computer on the Domain?", vbYesNo + vbQuestion, "Send Message to Everyone")
            Select Case Response
                Case 6
                Case 7
                    OptSingle.Value = True
                    Exit Sub
        End Select

        SendAll
    Else
        SendSingle
    End If
End Sub

Private Sub Form_Load()
    sndPlaySound App.Path & "\intro.wav", 1
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    DomainName = StrDomain
    UserDomain = StrDomain
    ContainerName = StrDomain
    Set Container = GetObject("WinNT://" & ContainerName)
    Set dso = GetObject("WinNT:")
    Container.Filter = Array("Computer")
    For Each Computer In Container
        Combo1.AddItem Computer.Name
    Next
    OptSingle.Value = True
End Sub

Public Function SendMessage(RcptToUser As String, _
   FromUser As String, BodyMessage As String) As Boolean

   Dim RcptTo() As Byte
   Dim From() As Byte
   Dim Body() As Byte

   RcptTo = RcptToUser & vbNullChar
   From = FromUser & vbNullChar
   Body = BodyMessage & vbNullChar

   If NetMessageBufferSend(ByVal 0&, RcptTo(0), ByVal 0&, _
        Body(0), UBound(Body)) = NERR_Success Then
     SendMessage = True
   End If
End Function

Private Sub EnableSend()
    Combo1.Enabled = True
    CmdSend.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    sndPlaySound App.Path & "\exit.wav", 0
    Unload Me
    Unload FrmGraphics
End Sub
Private Sub OptAll_Click()
    Dim Response As Long
    Response = MsgBox("Are you sure you want to send this message to Every computer on the Domain?", vbYesNo + vbQuestion, "Send Message to Everyone")
    Select Case Response
        Case 6
        Case 7
            OptSingle.Value = True
    End Select
    EnableSend
End Sub
Private Sub OptSingle_Click()
    EnableSend
End Sub
Private Sub GetFullName()
    Container.Filter = Array("User")
    Dim User As IADsUser
    For Each User In Container
        If User.Name = SysInfo.UserName Then
            StrUser = User.FullName
        End If
    Next
End Sub
Private Sub Timer_Timer()
    If IntFlame > 32 Then IntFlame = 0
    AniGif.Picture = FrmGraphics.flame(IntFlame).Picture
    IntFlame = IntFlame + 1
End Sub

