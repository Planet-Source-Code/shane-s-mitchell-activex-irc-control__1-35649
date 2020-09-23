VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ctlIRC 
   CanGetFocus     =   0   'False
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   2565
   ToolboxBitmap   =   "ctlIRC.ctx":0000
   Begin MSWinsockLib.Winsock ws 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image img 
      Height          =   720
      Left            =   0
      Picture         =   "ctlIRC.ctx":0312
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "ctlIRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************
'                        IRC EVENT-BASED CONTROL
' Author: Shane Mitchell (Extadi productions) http://www.extadi.com
' Authors Age: 13 (and 10 months)
' Last updated: 9th June 2002
'
' Please vote for me on Planet Source Code (www.pscode.com)
' And thanks for showing your interest
' Email me at founder@extadi.com
'**************************************************************************

Public Channels As New clnChannels

Public Type User
    Nick As String
    Ident As String
    Host As String
End Type

Public Enum ERR_Consts ' // this was the most anoying part of the whole thing, sheesh!
    RPL_UMODEIS = 221
    RPL_AWAY = 301
    RPL_CHANNELMODEIS = 324
    RPL_NOTOPIC = 331
    RPL_TOPIC = 332
    RPL_INVITING = 341
    RPL_NAMREPLY = 353
    RPL_ENDOFNAMES = 366
    RPL_BANLIST = 367
    RPL_ENDOFBANLIST = 368
    RPL_MOTDSTART = 375
    RPL_MOTD = 372
    RPL_ENDOFMOTD = 376
    ERR_NOSUCHNICK = 401
    ERR_NOSUCHCHANNEL = 403
    ERR_CANNOTSENDTOCHAN = 404
    ERR_TOOMANYCHANNELS = 405
    ERR_TOOMANYTARGETS = 407
    ERR_NORECIPIENT = 411
    ERR_NOTEXTTOSEND = 412
    ERR_NOTOPLEVEL = 413
    ERR_WILDTOPLEVEL = 414
    ERR_NOMOTD = 422
    ERR_NONICKNAMEGIVEN = 431
    ERR_ERRONEUSNICKNAME = 432
    ERR_NICKNAMEINUSE = 433
    ERR_NICKCOLLISION = 436
    ERR_NOTONCHANNEL = 442
    ERR_USERONCHANNEL = 443
    ERR_NEEDMOREPARAMS = 461
    ERR_KEYSET = 467
    ERR_CHANNELISFULL = 471
    ERR_UNKNOWNMODE = 472
    ERR_INVITEONLYCHAN = 473
    ERR_BANNEDFROMCHAN = 474
    ERR_BADCHANNELKEY = 475
    ERR_BADCHANMASK = 476
    ERR_CHANOPRIVSNEEDED = 482
    ERR_UMODEUNKNOWNFLAG = 501
    ERR_USERSDONTMATCH = 502

    RPL_WHOISUSER = 311
    RPL_WHOISOPERATOR = 313
    RPL_WHOISSERVER = 312
    RPL_WHOISIDLE = 317
    RPL_ENDOFWHOIS = 318
    RPL_WHOISCHANNELS = 319
    ERR_NOSUCHSERVER = 402

End Enum

Event ConnectionAuthing()
Event Connected()
Event Disconnect(ByVal sError As String)
Event NumEvent(ID As ERR_Consts, ByVal Description As String, ByVal Param1 As String, ByVal Param2 As String)
Event OnNotice(sFrom As User, ByVal sTo As String, ByVal sText As String)
Event OnPrivMsg(sFrom As User, ByVal sTo As String, ByVal sText As String)
Event OnJoin(sUser As User, ByVal sChannel As String)
Event OnPart(sUser As User, ByVal sChannel As String, ByVal sMessage As String)
Event OnQuit(sUser As User, ByVal sMessage As String)
Event OnInvite(sUser As User, ByVal sChannel As String)
Event OnTopic(sUser As User, ByVal sChannel As String, ByVal sTopic As String)
Event OnChannelMode(sUser As User, sChannel As String, sMode As String)
Event OnNick(sUser As User, ByVal sNewNick As String)
Event OnKick(sUser As User, ByVal sChannel As String, ByVal sKicked As String, sReason As String)
Event Raw(ByVal sData As String)

Private Const endL As String = " :"

Private State As enuStates
Private PVT_Nick As String
Private LocalUser As User

Public AUTH_USER As String
Public AUTH_NAME As String

Private Enum enuStates
    Closed = 0
    Connecting = 1
    Authing = 2
    Connected = 3
End Enum

Private Sub UserControl_Resize()
    UserControl.Width = img.Width
    UserControl.Height = img.Height
End Sub

Public Sub SetUser(sNick As String, sUser As String, sName As String)
    AUTH_USER = Replace(sUser, " ", "")
    AUTH_NAME = sName
    PVT_Nick = sNick
End Sub

Public Property Let Nick(sNick As String)
    SendData "NICK", sNick
End Property

Public Property Get Nick() As String
    Nick = PVT_Nick
End Property

Public Sub Connect(sServer As String, Optional sPort As Single = 6667)
    If AUTH_USER = "" Or AUTH_NAME = "" Or Nick = "" Then Err.Raise 10, , "Must SetUser() first"
    If State <> Closed Then
        Disconnect
    End If
    
    ws.Connect sServer, sPort
    State = Connecting
End Sub

Public Sub Disconnect(Optional sError As String = "")
    If State = Closed Then Exit Sub
    ws.Close
    
    RaiseEvent Disconnect(sError)
End Sub

Private Sub ws_Connect()
    RaiseEvent ConnectionAuthing
    
    SendData "USER", AUTH_USER, "server", "name", endL, AUTH_NAME
    SendData "NICK", PVT_Nick
End Sub

Private Sub SendData(ParamArray sPars() As Variant)
    On Error GoTo 1
    
    Dim i As Integer
    Dim sDta As String
    
    For i = LBound(sPars()) To UBound(sPars())
        sDta = sDta & sPars(i)
        If i <> UBound(sPars()) Then
            If sPars(i + 1) <> endL And sPars(i) <> endL Then
                sDta = sDta & " "
            End If
        End If
    Next i
    
    ws.SendData sDta & vbCrLf
    Exit Sub
1
    Disconnect "Write Error: Broken Pipe"
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo 1
    
    Static sOld As String
    Dim sTmp As String
    Dim a As Integer
    ws.GetData sTmp
    sTmp = sOld & sTmp
    sTmp = Replace(sTmp, vbCrLf, vbCr)
    sTmp = Replace(sTmp, vbLf, vbCr)
    sTmp = Replace(sTmp, vbCr & vbCr, vbCr)
    
    sOld = ""
    Do
        a = InStr(1, sTmp, vbCr)
        If a > 0 Then
            ProcessData Left(sTmp, a - 1)
            RaiseEvent Raw(Left(sTmp, a - 1))
            sTmp = Right(sTmp, Len(sTmp) - a)
        Else
            sOld = sTmp
            Exit Sub
        End If
    Loop Until sTmp = ""
    Exit Sub
1
    Disconnect "Read Error: Broken Pipe"
End Sub

Private Sub ProcessData(ByVal text As String)
    On Error Resume Next
    
    Dim sSmall() As String, sSpl() As String, sLarge As String
    Dim a As Integer, i As Integer, u As Integer
    Dim Mode As String, sTmp As String, sIN As String
    Dim sModes As String, sParams As String
    Dim bOnOff As Boolean, sPart As String * 1
    Dim Error As ERR_Consts
    Dim theUser As User
    
    If Trim(text) = "" Then Exit Sub
    If Left(text, 1) = ":" Then
        text = Right(text, Len(text) - 1)
        a = InStr(1, text, endL)
        If a > 0 Then sLarge = Right(text, Len(text) - a - 1): text = Left(text, a - 1)
        sSmall() = Split(text, " ")
        If UBound(sSmall()) < 0 Then Exit Sub
        
        If UBound(sSmall()) >= 1 Then
            If IsNumeric(sSmall(1)) Then ' // On Error or Reply
                Error = Val(sSmall(1))
                RaiseEvent NumEvent(Error, GetDesc(Error), sSmall(2), sLarge)
                
                If Error = RPL_ENDOFMOTD Or Error = ERR_NOMOTD Then
                    State = Connected
                    RaiseEvent Connected
                ElseIf Error = RPL_NAMREPLY Then
                    sSpl() = Split(sLarge, " ")
                    For i = 0 To UBound(sSpl())
                        Select Case Left(sSpl(i), 1)
                        Case "@", "+"
                            Mode = Left(sSpl(i), 1)
                            sSpl(i) = Right(sSpl(i), Len(sSpl(i)) - 1)
                        Case Else
                            Mode = ""
                        End Select
                        Channels(sSmall(UBound(sSmall()))).Nicks.Add sSpl(i), sSpl(i), "", "", Mode, sSpl(i)
                    Next i
                ElseIf Error = RPL_TOPIC Then
                    Channels(sSmall(UBound(sSmall()))).Topic = sLarge
                ElseIf Error = RPL_ENDOFNAMES Then
                    RaiseEvent OnJoin(LocalUser, sSmall(UBound(sSmall())))
                End If
            Else
                theUser = GetUser(sSmall(0))
                If LCase(theUser.Nick) = LCase(PVT_Nick) Then
                    LocalUser = theUser
                End If
                Select Case UCase(sSmall(1))
                Case "NOTICE"
                    RaiseEvent OnNotice(theUser, sSmall(2), sLarge)
                Case "PRIVMSG"
                    RaiseEvent OnPrivMsg(theUser, sSmall(2), sLarge)
                Case "JOIN"
                    If sLarge = "" Then
                        If UBound(sSmall()) > 1 Then
                            sTmp = sSmall(UBound(sSmall()))
                        End If
                    Else
                        sTmp = sLarge
                    End If
                    If LCase(theUser.Nick) = LCase(Nick) Then
                        Channels.Add sTmp, sTmp, "", sTmp
                    Else
                        Channels(sTmp).Nicks.Add theUser.Nick, theUser.Nick, theUser.Ident, theUser.Host, "", theUser.Nick
                        RaiseEvent OnJoin(theUser, sTmp)
                    End If
                Case "PART"
                    If LCase(theUser.Nick) = LCase(Nick) Then
                        Channels.Remove sSmall(2)
                        RaiseEvent OnPart(LocalUser, sSmall(2), sLarge)
                    Else
                        Channels(sSmall(2)).Nicks.Remove theUser.Nick
                        RaiseEvent OnPart(theUser, sSmall(2), sLarge)
                    End If
                Case "QUIT"
                    For i = 1 To Channels.Count
                        For u = 1 To Channels(i).Nicks.Count
                            If LCase(Channels(i).Nicks(u).Nick) = LCase(theUser.Nick) Then
                                Channels(i).Nicks.Remove u
                                Exit For
                            End If
                        Next u
                    Next i
                    RaiseEvent OnQuit(theUser, sLarge)
                Case "INVITE"
                    If UBound(sSmall()) >= 3 Then
                        RaiseEvent OnInvite(theUser, sSmall(3))
                    Else
                        RaiseEvent OnInvite(theUser, sLarge)
                    End If
                Case "TOPIC"
                    Channels(sSmall(2)).Topic = sLarge
                    RaiseEvent OnTopic(theUser, sSmall(2), sLarge)
                Case "MODE"
                    If LCase(sSmall(2)) <> LCase(Nick) Then
                        If Trim(sLarge) <> "" Then
                            sIN = sLarge
                        Else
                            For i = 3 To UBound(sSmall())
                                sTmp = sTmp & IIf((Len(sTmp) > 0), " ", "") & sSmall(i)
                            Next i
                            sIN = sTmp
                        End If
                        a = InStr(1, sIN, " ")
                        If a > 0 Then
                            sModes = Trim(Left(sIN, a - 1))
                            sParams = Trim(Right(sIN, Len(sIN) - a))
                            sSpl() = Split(sParams, " ")
                            a = 0
                            bOnOff = True
                            
                            For i = 1 To Len(sModes)
                                sPart = LCase(Mid(sModes, i, 1))
                                If sPart = "+" Then
                                    bOnOff = True
                                ElseIf sPart = "-" Then
                                    bOnOff = False
                                ElseIf sPart = "o" Or sPart = "b" Or sPart = "v" Or sPart = "k" Then
                                    If sPart = "o" Or sPart = "v" Then
                                        Mode = Channels(sSmall(2)).Nicks(sSpl(a)).Mode
                                        Mode = Replace(Mode, sPart, "")
                                        If bOnOff Then
                                            Mode = Mode & sPart
                                        End If
                                        Channels(sSmall(2)).Nicks(sSpl(a)).Mode = Mode
                                    ElseIf sPart = "k" Then
                                        If bOnOff Then
                                            Channels(sSmall(2)).Password = sSpl(a)
                                        Else
                                            Channels(sSmall(2)).Password = ""
                                        End If
                                    ElseIf sPart = "b" Then
                                        For u = 1 To Channels(sSmall(2)).Bans.Count
                                            If LCase(Channels(sSmall(2)).Bans(u)) = LCase(sSpl(a)) Then
                                                Channels(sSmall(2)).Bans.Remove u
                                            End If
                                        Next u
                                        If bOnOff Then
                                            Channels(sSmall(2)).Bans.Add sSpl(a), sSpl(a), sSpl(a)
                                        End If
                                    Else
                                        Channels(sSmall(2)).Modes = Replace(Channels(sSmall(2)).Modes, sPart, "")
                                        If bOnOff Then
                                            Channels(sSmall(2)).Modes = Channels(sSmall(2)).Modes & sPart
                                        End If
                                    End If
                                    RaiseEvent OnChannelMode(theUser, sSmall(2), IIf(bOnOff, "+", "-") & sPart & " " & sSpl(a))
                                    a = a + 1
                                Else
                                    RaiseEvent OnChannelMode(theUser, sSmall(2), IIf(bOnOff, "+", "-") & sPart)
                                End If
                            Next i
                        End If
                    End If
                Case "NICK"
                    Dim sHOLDER(0 To 2) As String
                    For i = 1 To Channels.Count
                        For u = 1 To Channels(i).Nicks.Count
                            If LCase(Channels(i).Nicks(u).Nick) = LCase(theUser.Nick) Then
                                sHOLDER(0) = Channels(i).Nicks(u).Host
                                sHOLDER(1) = Channels(i).Nicks(u).Ident
                                sHOLDER(2) = Channels(i).Nicks(u).Mode
                                Channels(i).Nicks.Remove u
                                Channels(i).Nicks.Add sLarge, sLarge, sHOLDER(1), sHOLDER(0), sHOLDER(2), sLarge
                                Exit For
                            End If
                        Next u
                    Next i
                    RaiseEvent OnNick(theUser, sLarge)
                Case "KICK"
                    If LCase(sSmall(3)) = LCase(Nick) Then ' damn thats us getting kicked
                        Channels.Remove sSmall(2)
                    Else
                        Channels(sSmall(2)).Nicks.Remove sSmall(3)
                    End If
                    RaiseEvent OnKick(theUser, sSmall(2), sSmall(3), sLarge)
                End Select
            End If
        End If
    Else
        a = InStr(1, text, endL)
        If a > 0 Then sLarge = Right(text, Len(text) - a - 1): text = Left(text, a - 1)
        sSmall() = Split(text, " ")
        If UBound(sSmall()) < 0 Then Exit Sub
        
        Select Case UCase(sSmall(0))
        Case "PING"
            SendData "PONG", endL, sLarge
        End Select
    End If
End Sub

Public Sub Raw(sData As String)
    On Error Resume Next
    ws.SendData sData & vbCrLf
End Sub

Public Sub JoinChans(ParamArray sChans() As Variant)  ' // as in: #dog #cat #freddo smackers
    Dim i As Integer
    Dim sChan As String, sPass As String
    For i = LBound(sChans()) To UBound(sChans())
        If Left(sChans(i), 1) = "#" Then
            sChan = sChan & IIf((Len(sChan) > 0), ",", "") & sChans(i)
        Else
            sPass = sPass & IIf((Len(sPass) > 0), ",", "") & sChans(i)
        End If
    Next i
    SendData "JOIN", sChan, sPass
End Sub

Public Sub PartChans(sMessage As String, ParamArray sChans() As Variant)
    Dim i As Integer
    Dim sChan As String
    For i = LBound(sChans()) To UBound(sChans())
        sChan = sChan & IIf((Len(sChan) > 0), ",", "") & sChans(i)
    Next i
    If sMessage <> "" Then
        SendData "PART", sChan, sMessage
    Else
        SendData "PART", sChan
    End If
End Sub

Public Sub PrivMsg(sText As String, ParamArray sPpl() As Variant)
    Dim i As Integer
    Dim sTmp As String
    For i = LBound(sPpl()) To UBound(sPpl())
        sTmp = sTmp & IIf((Len(sTmp) > 0), ",", "") & sPpl(i)
    Next i
    SendData "PRIVMSG", sTmp, endL, sText
End Sub

Public Sub Quit(sText As String)
    SendData "QUIT", endL, sText
End Sub

Public Sub Mode(sRecipient As String, ParamArray sModes() As Variant)
    Dim sTmp As String, sTmp2 As String, sTmp3, sPart As String, sSpl() As String
    Dim i As Integer
    For i = LBound(sModes()) To UBound(sModes())
        sPart = Trim(sModes(i))
        sSpl() = Split(sPart, " ")
        If UBound(sSpl()) = 0 Then
            sTmp2 = sTmp2 + sPart
        Else
            sTmp = sTmp & sSpl(0)
            sTmp3 = sTmp3 & IIf((Len(sTmp3) > 0), " ", "") & sSpl(1)
        End If
    Next i
    
    SendData "MODE", sRecipient, sTmp & sTmp2, sTmp3
End Sub

Public Sub Topic(sChan As String, sTopic As String)
    SendData "TOPIC", sChan, endL, sTopic
End Sub

Public Sub Invite(sChan As String, sNick As String)
    SendData "INVITE", sNick, sChan
End Sub

Public Sub Kick(sChan As String, sNick As String, Optional sMessage As String = "")
    If sMessage = "" Then sMessage = PVT_Nick
    SendData "KICK", sChan, sNick, endL, sMessage
End Sub

Public Sub Notice(sText As String, ParamArray sPpl() As Variant)
    Dim i As Integer
    Dim sTmp As String
    For i = LBound(sPpl()) To UBound(sPpl())
        sTmp = sTmp & IIf((Len(sTmp) > 0), ",", "") & sPpl(i)
    Next i
    
    SendData "NOTICE", sTmp, endL, sText
End Sub

Public Sub WhoIs(sNick As String)
    SendData "WHOIS", sNick
End Sub

Private Function GetUser(sText As String) As User
    Dim sTmp() As String
    sTmp() = Split(sText, "@")
    If UBound(sTmp()) = 1 Then
        GetUser.Host = sTmp(1)
        sTmp() = Split(sTmp(0), "!")
        If UBound(sTmp()) = 1 Then
            GetUser.Ident = sTmp(1)
        End If
        GetUser.Nick = sTmp(0)
    Else
        sTmp() = Split(sText, "!")
        If UBound(sTmp()) = 1 Then
            GetUser.Ident = sTmp(1)
        End If
        GetUser.Nick = sTmp(0)
    End If
End Function

Public Function GetDesc(iError As ERR_Consts) As String
    Select Case iError
    Case RPL_AWAY
        GetDesc = "Away"
    Case RPL_NOTOPIC
        GetDesc = "No topic set"
    Case RPL_TOPIC
        GetDesc = "Topic"
    Case RPL_MOTDSTART
        GetDesc = "Beginning of MOTD"
    Case RPL_MOTD
        GetDesc = "MOTD"
    Case RPL_ENDOFMOTD
        GetDesc = "End of MOTD"
    Case ERR_NOSUCHNICK
        GetDesc = "No such nick"
    Case ERR_NOSUCHCHANNEL
        GetDesc = "No such Channel"
    Case ERR_CANNOTSENDTOCHAN
        GetDesc = "Cannot send to channel"
    Case ERR_TOOMANYCHANNELS
        GetDesc = "You are in too many channels"
    Case ERR_TOOMANYTARGETS
        GetDesc = "Too many targets"
    Case ERR_NORECIPIENT
        GetDesc = "No recipient"
    Case ERR_NOTEXTTOSEND
        GetDesc = "No text to send"
    Case ERR_NOTOPLEVEL
        GetDesc = "No top level"
    Case ERR_WILDTOPLEVEL
        GetDesc = "Wild top level"
    Case ERR_NOMOTD
        GetDesc = "MOTD file is missing"
    Case ERR_NONICKNAMEGIVEN
        GetDesc = "No nickname given"
    Case ERR_ERRONEUSNICKNAME
        GetDesc = "Erroneus Nickname"
    Case ERR_NICKNAMEINUSE
        GetDesc = "Nickname already in use"
    Case ERR_NICKCOLLISION
        GetDesc = "Nick collision"
    Case ERR_NOTONCHANNEL
        GetDesc = "Not on channel"
    Case ERR_NEEDMOREPARAMS
        GetDesc = "Not enough paramaters"
    Case ERR_CHANNELISFULL
        GetDesc = "Channel is full"
    Case ERR_INVITEONLYCHAN
        GetDesc = "Invite Only"
    Case ERR_BANNEDFROMCHAN
        GetDesc = "You are banned from the channel"
    Case ERR_BADCHANNELKEY
        GetDesc = "Invalid key"
    Case ERR_BADCHANMASK
        GetDesc = "Bad channel mask"
    Case RPL_UMODEIS
        GetDesc = "User Mode"
    Case RPL_CHANNELMODEIS
        GetDesc = "Channel Mode"
    Case RPL_BANLIST
        GetDesc = "Ban list"
    Case RPL_ENDOFBANLIST
        GetDesc = "End of Ban list"
    Case ERR_KEYSET
        GetDesc = "Key already set"
    Case ERR_UNKNOWNMODE
        GetDesc = "Unknown Mode"
    Case ERR_CHANOPRIVSNEEDED
        GetDesc = "You are not a channel operator"
    Case ERR_UMODEUNKNOWNFLAG
        GetDesc = "Unknown MODE flag"
    Case RPL_INVITING
        GetDesc = "Inviting"
    Case ERR_USERONCHANNEL
        GetDesc = "User already on channel"
    Case RPL_NAMREPLY
        GetDesc = "Nick List"
    Case RPL_ENDOFNAMES
        GetDesc = "End of Names List"
    End Select
End Function

Public Function IsOp(sNick As String, sChannel As String) As Boolean
    On Error GoTo 1
    If InStr(1, Channels(sChannel).Nicks(sNick).Mode, "o") > 0 Then
        IsOp = True
    End If
1 End Function

Public Function IsVoice(sNick As String, sChannel As String) As Boolean
    On Error GoTo 1
    If InStr(1, Channels(sChannel).Nicks(sNick).Mode, "v") > 0 Then
        IsVoice = True
    End If
1 End Function

Public Function IsOn(sNick As String, sChannel As String) As Boolean ' determines whether someone is on a channel
    On Error GoTo 1
    Dim sDummy As String
    sDummy = Channels(sChannel).Nicks(sNick).Nick ' should return an error if their not on the channel and skip to 1
    IsOn = True
1 End Function
