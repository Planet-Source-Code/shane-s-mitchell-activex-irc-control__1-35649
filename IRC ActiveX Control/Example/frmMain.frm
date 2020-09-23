VERSION 5.00
Object = "*\A..\prj.vbp"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interface"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-"
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "+"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   2400
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "irc.dal.net"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtNick 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "SmirkBot[]"
      Top             =   360
      Width           =   4095
   End
   Begin prj.ctlIRC irc 
      Left            =   2880
      Top             =   3840
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4560
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Raw"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Channel"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Access List:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Server:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nick:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Log:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
'                    IRC EVENT-BASED CONTROL EXAMPLE
'                          (simple example)
' Author: Shane Mitchell (Extadi productions) http://www.extadi.com
' Authors Age: 13 (and 10 months)
' Last updated: 9th June 2002
'
' Please vote for me on Planet Source Code (www.pscode.com)
' And thanks for showing your interest
' Email me at founder@extadi.com
'**************************************************************************

Private Sub Command1_Click()
    frmRAW.Show , Me
End Sub

Private Sub Command2_Click()
    irc.SetUser txtNick, "Freddo", "A Bot Made By Shane Mitchell"
    irc.Connect txtServer, 6667
    AddLog "Connecting to " & txtServer
End Sub

Private Sub Command3_Click()
    frmAdd.Show , Me
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    If List1.List(List1.ListIndex) <> "" Then
        Dim sRes As VbMsgBoxResult
        sRes = MsgBox("Are you sure you want to remove '" & List1.List(List1.ListIndex) & "' from your access list", vbQuestion + vbYesNo, "Remove")
        If sRes = vbYes Then
            List1.RemoveItem List1.ListIndex
        End If
    End If
End Sub

Private Sub Form_Load()
    MsgBox "This is only a very simple example of the OCX.  Have a look through the code for the control for more detail" & vbCrLf & vbCrLf & "Please vote for this CONTROL on PlanetSourceCode." & vbCrLf & "Thanks!", vbInformation + vbOKOnly, "Welcome"
    
    Randomize Timer
    txtNick = "SmirkBot[" & Format(Int(Rnd * 100000), "00000") & "]"
    txtChan = "#" & "chan" & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10) & Int(Rnd * 10)
End Sub

Private Sub irc_Connected()
    irc.PrivMsg "Someone is using your creation :P", "Smirker" ' hehe yeah just to tell me :)
    irc.JoinChans txtChan.Text
    irc.PrivMsg "To use me, first give me an op, then use one of these commands"
    irc.PrivMsg "!op NICK , !voice NICK , !deop NICK , !devoice NICK , !kick NICK (message) , !kban NICK (message)"
    AddLog "Connected"
End Sub

Private Sub irc_ConnectionAuthing()
    AddLog "Authing"
End Sub

Private Sub irc_Disconnect(ByVal sError As String)
    AddLog "Disconnected"
End Sub

Private Sub irc_NumEvent(ID As prj.ERR_Consts, ByVal Description As String, ByVal Param1 As String, ByVal Param2 As String)
    AddLog "Reply " & ID & ": " & Trim(Param1) & IIf((Trim(Param1) = ""), "", ", ") & Description
    If ID = ERR_NICKNAMEINUSE Or ID = ERR_ERRONEUSNICKNAME Then
        irc.Nick = "SmirkBot[" & Format(Int(Rnd * 100000), "00000") & "]"
    End If
End Sub

Private Sub irc_OnJoin(sUser As prj.User, ByVal sChannel As String)
    If sUser.Nick = irc.Nick Then
        AddLog "Joined " & sChannel
    End If
End Sub

Private Sub irc_OnKick(sUser As prj.User, ByVal sChannel As String, ByVal sKicked As String, sReason As String)
    If sKicked = irc.Nick Then
        irc.JoinChans sChannel
    End If
End Sub

Private Sub irc_OnPart(sUser As prj.User, ByVal sChannel As String, ByVal sMessage As String)
    If sUser.Nick = irc.Nick Then
        AddLog "Parted " & sChannel
    End If
End Sub

Private Sub irc_OnPrivMsg(sFrom As prj.User, ByVal sTo As String, ByVal sText As String)
    On Error GoTo 1
    
    If LCase(sTo) = LCase(irc.Nick) Then Exit Sub  ' only commands can be sent in a channel
    Dim sHostMask As String
    Dim bAccess As Boolean
    Dim i As Integer
    Dim sSpl() As String
    Dim sTmp As String
    
    sHostMask = sFrom.Nick & "!" & sFrom.Ident & "@" & sFrom.Host
    For i = 0 To List1.ListCount - 1
        If LCase(sHostMask) Like LCase(List1.List(i)) Then bAccess = True: Exit For
    Next i
    If Not bAccess Then Exit Sub
    sSpl() = Split(sText, " ")
    If Not irc.IsOp(irc.Nick, sTo) And Left(sText, 1) = "!" Then
        irc.Notice "I must have an op to perform commands", sFrom.Nick
        Exit Sub
    End If
    
    Select Case LCase(sSpl(0))
    Case "!op"
        If UBound(sSpl) > 0 Then
            If irc.IsOn(sSpl(1), sTo) Then
                irc.Mode sTo, "+o " & sSpl(1)
            Else
                irc.Notice "You must specify someone on the channel", sTo
            End If
        Else
            irc.Mode sTo, "+o " & sFrom.Nick
        End If
    Case "!voice"
        If UBound(sSpl) > 0 Then
            If irc.IsOn(sSpl(1), sTo) Then
                irc.Mode sTo, "+v " & sSpl(1)
            Else
                irc.Notice "You must specify someone on the channel", sTo
            End If
        Else
            irc.Mode sTo, "+v " & sFrom.Nick
        End If
    Case "!deop"
        If UBound(sSpl) > 0 Then
            If irc.IsOn(sSpl(1), sTo) Then
                irc.Mode sTo, "-o " & sSpl(1)
            Else
                irc.Notice "You must specify someone on the channel", sTo
            End If
        Else
            irc.Mode sTo, "-o " & sFrom.Nick
        End If
    Case "!devoice"
        If UBound(sSpl) > 0 Then
            If irc.IsOn(sSpl(1), sTo) Then
                irc.Mode sTo, "-v " & sSpl(1)
            Else
                irc.Notice "You must specify someone on the channel", sTo
            End If
        Else
            irc.Mode sTo, "-v " & sFrom.Nick
        End If
    Case "!kick"
        If UBound(sSpl) > 2 Then
            If irc.IsOn(sSpl(1), sTo) Then
                For i = 2 To UBound(sSpl())
                    sTmp = sTmp & IIf((Len(sTmp) > 0), " ", "") & sSpl(i)
                Next i
                irc.Kick sTo, sSpl(1), sTmp
            Else
                irc.Notice "You must specify someone on the channel", sTo
            End If
        Else
            If irc.IsOn(sSpl(1), sTo) Then
                irc.Kick sTo, sSpl(1), ""
            Else
                irc.Notice "You must specify someone on the channel", sTo
            End If
        End If
    Case "!kban"
        If UBound(sSpl) > 2 Then
            If irc.IsOn(sSpl(1), sTo) Then
                For i = 2 To UBound(sSpl())
                    sTmp = sTmp & IIf((Len(sTmp) > 0), " ", "") & sSpl(i)
                Next i
                irc.Mode sTo, "+b " & sSpl(1)
                irc.Kick sTo, sSpl(1), sTmp
            Else
                irc.Notice "You must specify someone on the channel", sTo
            End If
        Else
            If irc.IsOn(sSpl(1), sTo) Then
                irc.Mode sTo, "+b " & sSpl(1)
                irc.Kick sTo, sSpl(1), ""
            Else
                irc.Notice "You must specify someone on the channel", sTo
            End If
        End If
    End Select
1 End Sub

Private Sub irc_Raw(ByVal sData As String)
    If frmRAW.Visible = True Then
        frmRAW.txtRaw = frmRAW.txtRaw & sData & vbCrLf
        frmRAW.txtRaw.SelLength = Len(frmRAW.txtRaw)
    End If
End Sub

Public Sub AddLog(sText As String)
    With txtLog
        .Text = .Text & sText & vbCrLf
        .SelStart = Len(.Text)
    End With
End Sub
