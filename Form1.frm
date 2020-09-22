VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckTCP 
      Index           =   0
      Left            =   4080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()
MaxSck = 20 'maximum number of connections
ServerName = "IRCS" 'set server name

ReDim Buffer(MaxSck)
ReDim Client(MaxSck)

For a = 1 To MaxSck
 Load sckTCP(a)
Next a

'size dynamic arrays
Call Size_Array

'start listening for connections
sckTCP(0).LocalPort = 6667
sckTCP(0).Listen

Me.Caption = "Maximum connections: " & MaxSck & " ~ " & "Servername: " & ServerName
End Sub

Private Sub sckTCP_Close(Index As Integer)
Call Socket_Close(Index)
End Sub

Private Sub sckTCP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
sckTCP(0).Close

For a = 1 To MaxSck

If sckTCP(a).State = 0 Then
 sckTCP(a).Accept requestID
 sckTCP(a).SendData ":" & ServerName & " NOTICE :Yay you're connected :)" & vbCrLf


'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 002 c6d :Your host is astro.ga.us.dal.net[@0.0.0.0], running version bahamut-1.4(34).rhashfix" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 003 c6d :This server was created Tue Jul 30 2002 at 19:15:22 EDT" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 004 c6d astro.ga.us.dal.net bahamut-1.4(34).rhashfix oiwscrknfydaAbghe biklLmMnoprRstvc" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 005 c6d NOQUIT WATCH=128 SAFELIST MODES=13 MAXCHANNELS=20 MAXBANS=100 NICKLEN=30 TOPICLEN=307 KICKLEN=307 CHANTYPES=&# PREFIX=(ov)@+ NETWORK=DALnet SILENCE=10 CASEMAPPING=ascii :are available on this server" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 251 c6d :There are 4659 users and 84502 invisible on 25 servers" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 252 c6d 69 :IRC Operators online" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 254 c6d 33733 :channels formed" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 255 c6d :I have 5709 clients and 1 servers" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 265 c6d :Current local users: 5709 Max: 25358" & vbCrLf
'Form1.sckTCP(a).SendData ":astro.ga.us.dal.net 266 c6d :Current global users: 89161 Max: 138996" & vbCrLf
 
 
 sckTCP(0).Listen
 Exit For
End If

Next a
End Sub



Private Sub sckTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
sckTCP(Index).GetData IncData

'Text1 = Text1 & IncData 'dump incoming data to text box

If Right(IncData, 1) <> Chr(10) Then
 Buffer(Index) = Buffer(Index) & IncData
Else
 Buffer(Index) = Buffer(Index) & IncData
 Call Process_IncData(Index)
End If

End Sub




Public Function holdthiscode()
If Left(IncData, 4) = "NICK" Then
 IncData = Mid(IncData, 6)
 ':astro.ga.us.dal.net NOTICE AUTH :*** Looking up your hostname...
 'MsgBox Asc(Mid(temp$, 4))
 sckTCP(Index).Tag = Left(temp$, 3)
 sckTCP(Index).SendData ":IRCS NOTICE :Yay you're connected :)" & vbCrLf
 'sckTCP(Index).SendData ":c6d!c6d.net@h00500460dc12.ne.client2.attbi.com JOIN :#IRCS" & vbCrLf
 sckTCP(Index).SendData ":c6d JOIN :#IRCS" & vbCrLf
 'sckTCP(Index).SendData ":astro.ga.us.dal.net NOTICE AUTH :*** Looking up your hostname..." & vbCrLf
 'MsgBox Len(sckTCP(Index).Tag)
End If

':astro.ga.us.dal.net NOTICE AUTH :*** Looking up your hostname...
':astro.ga.us.dal.net NOTICE AUTH :*** Found your hostname, cached
':astro.ga.us.dal.net NOTICE AUTH :*** Checking Ident
':astro.ga.us.dal.net NOTICE AUTH :*** Got Ident response
':astro.ga.us.dal.net 001 Bot_11 :Welcome to the DALnet IRC Network Bot_11!c6d.net@h00500460dc12.ne.client2.attbi.com
':astro.ga.us.dal.net 002 Bot_11 :Your host is astro.ga.us.dal.net[@0.0.0.0], running version bahamut-1.4(34).rhashfix
'NOTICE Bot_11 :*** Your host is astro.ga.us.dal.net[@0.0.0.0], running version bahamut-1.4(34).rhashfix
':astro.ga.us.dal.net 003 Bot_11 :This server was created Tue Jul 30 2002 at 19:15:22 EDT
':astro.ga.us.dal.net 004 Bot_11 astro.ga.us.dal.net bahamut-1.4(34).rhashfix oiwscrknfydaAbghe biklLmMnoprRstvc
':astro.ga.us.dal.net 005 Bot_11 NOQUIT WATCH=128 SAFELIST MODES=13 MAXCHANNELS=20 MAXBANS=100 NICKLEN=30 TOPICLEN=307 KICKLEN=307 CHANTYPES=&# PREFIX=(ov)@+ NETWORK=DALnet SILENCE=10 CASEMAPPING=ascii :are available on this server
':astro.ga.us.dal.net 251 Bot_11 :There are 4659 users and 84502 invisible on 25 servers
':astro.ga.us.dal.net 252 Bot_11 69 :IRC Operators online
':astro.ga.us.dal.net 254 Bot_11 33733 :channels formed
':astro.ga.us.dal.net 255 Bot_11 :I have 5709 clients and 1 servers
':astro.ga.us.dal.net 265 Bot_11 :Current local users: 5709 Max: 25358
':astro.ga.us.dal.net 266 Bot_11 :Current global users: 89161 Max: 138996
':astro.ga.us.dal.net NOTICE Bot_11 :*** Notice -- motd was last changed at 10/7/2002 22:45
':astro.ga.us.dal.net NOTICE Bot_11 :*** Notice -- Please read the motd if you haven't read it
':astro.ga.us.dal.net 375 Bot_11 :- astro.ga.us.dal.net Message of the Day -
':astro.ga.us.dal.net 372 Bot_11 :- astro.ga.us.dal.net, running FreeBSD on a Athlon 900.
':astro.ga.us.dal.net 372 Bot_11 :-                      linked since December 8, 2000
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :-         01100001 01110011 01110100 01110010 01101111
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- ---- Disclaimer ---------------------------------------------------------
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- By connecting to this server you agree to be bound by the terms put forth
':astro.ga.us.dal.net 372 Bot_11 :- in DALnet's Acceptable Use Policy (http://www.dal.net/aup/)
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- ---- Server Staff -------------------------------------------------------
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- admin: next (next@dal.net), tbaur (tbaur@dal.net)
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- chief spanking ninja: eck
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- opers: linder, erho, bskin, adm, eck, pell (vanity-staff@dal.net)
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- ---- Server Rules -------------------------------------------------------
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- requests to be made an irc operator will be ignored, so just don't do it.
':astro.ga.us.dal.net 372 Bot_11 :- if you have to think about whether what you are doing is wrong, it is.
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- we reserve the right to do whatever we want to users on this server. that
':astro.ga.us.dal.net 372 Bot_11 :- includes removing access without warning if you are abusive, and sending
':astro.ga.us.dal.net 372 Bot_11 :- you to jail if you are a packet kiddie.
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :-         enjoy!
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- ---- Other --------------------------------------------------------------
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- want to show your appreciation to the folks who spend their time running
':astro.ga.us.dal.net 372 Bot_11 :- this server? send them a cd on their wishlist!
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :-         http://www.panaso.com/tbaur/wishlist
':astro.ga.us.dal.net 372 Bot_11 :-         http://arpa.com/next/wishlist
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- ---- Sponsor ------------------------------------------------------------
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :- This server graciously hosted by webusenet (http://www.webusenet.com)
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 372 Bot_11 :-         ANONYMITY, COMPLETION, SPEED, RETENTION
':astro.ga.us.dal.net 372 Bot_11 :-         NEWS FEEDS FROM WWW.USENETSERVER.COM
':astro.ga.us.dal.net 372 Bot_11 :-         CURRENTLY CARRYING OVER 31 THOUSAND GROUPS
':astro.ga.us.dal.net 372 Bot_11 :-
':astro.ga.us.dal.net 376 Bot_11 :End of /MOTD command.
':astro.ga.us.dal.net NOTICE Bot_11 :*** Notice -- This server runs an open proxy monitor to prevent abuse.
':astro.ga.us.dal.net NOTICE Bot_11 :*** Notice -- If you see connections on various ports from proxy.monitor.dal.net
':astro.ga.us.dal.net NOTICE Bot_11 :*** Notice -- please disregard them, as they are the monitor in action.
':astro.ga.us.dal.net NOTICE Bot_11 :*** Notice -- For more information please visit http://www.dal.net/proxies/



End Function

