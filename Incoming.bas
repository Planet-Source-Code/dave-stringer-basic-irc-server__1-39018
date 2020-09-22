Attribute VB_Name = "Incoming"
Public Function Process_IncData(Index As Integer)

'strip chr 13 if it exists
Buffer(Index) = Replace(Buffer(Index), Chr(13), "") 'vbcr chr(13)

'split the data by line feeds chr 10
The_Data = Split(Buffer(Index), Chr(10))

For a = 0 To UBound(The_Data) - 1

 If InStr(1, The_Data(a), " ") = 0 Then
  The_Data(a) = The_Data(a) & " " 'handles 1 word splits
 End If
 
 Select Case UCase(Left(The_Data(a), InStr(1, The_Data(a), " ") - 1))
 
  Case Is = "NICK"
   Call Nick(Index, The_Data(a))
   Buffer(Index) = ""
  Case Is = "JOIN"
   Call JOINc(Index, The_Data(a))
   Buffer(Index) = ""
  Case Is = "PART"
   Call PART(Index, The_Data(a))
   Buffer(Index) = ""
  Case Is = "PRIVMSG"
   Call PRIVMSG(Index, The_Data(a))
   Buffer(Index) = ""
  Case Is = "MODE"
   Call MODE(Index, The_Data(a))
   Buffer(Index) = ""
  Case Is = "NOTICE"
   Call NOTICE(Index, The_Data(a))
   Buffer(Index) = ""
 End Select


Next a
End Function

Public Function Nick(Index As Integer, The_Data)
'assign nick to socket if not there. means it is a new connection
If Form1.sckTCP(Index).Tag = "" Then
 'assign nick to socket tag
 Form1.sckTCP(Index).Tag = Mid(The_Data, 6)
 'add data to client info
 Client(Index) = Mid(The_Data, 6) & ":"
'some irc clients need this stuff
 Form1.sckTCP(Index).SendData ":" & ServerName & " 001 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 002 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 003 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 004 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 005 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 251 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 252 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 254 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 255 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 265 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 266 " & Form1.sckTCP(Index).Tag & " :" & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 372 " & Form1.sckTCP(Index).Tag & " :The MOTD." & vbCrLf
 Form1.sckTCP(Index).SendData ":" & ServerName & " 376 " & Form1.sckTCP(Index).Tag & " :End of the MOTD" & vbCrLf
'end of the crap :)
End If
End Function

Public Function JOINc(Index As Integer, The_Data)
Dim cName As String 'channel name
Dim cExist As Boolean 'does the channel exist?
Dim cEmpty As Boolean 'use a spot in the array that is empty?
Dim cIndex 'track channel array index if known to make things speedy

cName = Mid(The_Data, 6) 'seperate channel
'cName = Left(cName, Len(cName) - 1) 'seperate channel

'handle joining multiple channels via /join #a,#b,#c,#d
If InStr(1, cName, ",") > 0 Then
 tmpString = Split(cName, ",")
 For a = 0 To UBound(tmpString)
  Call JOINc(Index, "JOIN " & tmpString(a))
 Next a
 Exit Function
End If

cExist = False
cEmpty = False

'does channel exist?
If UBound(Channel) > 0 Then 'don't bother if channels don't exist
 For a = 1 To UBound(Channel)
  
  If Channel(a) = "" Then 'empty spot in the array so use it!
   cEmpty = True
   cIndex = a 'track what index the existing array was found at
   Exit For
  End If
  
  If cName = Left(Channel(a), InStr(1, Channel(a), ":") - 1) Then
   cExist = True
   cIndex = a 'track what index the existing channel was found at
   Exit For
  End If
 
 Next a
End If

'channel does not exist
'channel(whatever) = #channel:+modes:nick:nick:nick:nick
If cExist = False Then
 
 If cEmpty = False Then 'add new channel to end of the array
  ReDim Preserve Channel(UBound(Channel) + 1) 'add 1 to the channel array
  Channel(UBound(Channel)) = cName & ":+:" & Form1.sckTCP(Index).Tag & ":" 'add channel and new nick to list
 Else 'use the empty spot in the array
  Channel(cIndex) = cName & ":+:" & Form1.sckTCP(Index).Tag & ":"
 End If
 Form1.sckTCP(Index).SendData ":" & Form1.sckTCP(Index).Tag & "!.@" & Form1.sckTCP(Index).RemoteHostIP & " JOIN :" & cName & vbCrLf
 ':astro.ga.us.dal.net 353 Bot_11 = #ai :Bot_11 @c6d @Bot1 @Kahless
 Form1.sckTCP(Index).SendData ":" & ServerName & " 353 " & Form1.sckTCP(Index).Tag & " = " & cName & " :" & Form1.sckTCP(Index).Tag & vbCrLf
 ':astro.ga.us.dal.net 366 Bot_11 #AI :End of /NAMES list.
 Form1.sckTCP(Index).SendData ":" & ServerName & " 366 " & Form1.sckTCP(Index).Tag & " " & cName & " :End of /NAMES list." & vbCrLf
 Client(Index) = Client(Index) & cName & ":"
 Exit Function
End If

'channel does exist
'client(whatever) = nick:#channel:#channel:#channel:
For a = 1 To UBound(Client)
 If InStr(1, Client(a), ":" & cName & ":") <> 0 Then
 ':Bot_11!c6d.net@h00500460dc12.ne.client2.attbi.com JOIN :#AI
  'send to clients in channel of new join
  Form1.sckTCP(a).SendData ":" & Form1.sckTCP(Index).Tag & "!.@" & Form1.sckTCP(Index).RemoteHostIP & " JOIN :" & cName & vbCrLf
 End If
Next a
 
'send to client that joined that he joined
Form1.sckTCP(Index).SendData ":" & Form1.sckTCP(Index).Tag & "!.@" & Form1.sckTCP(Index).RemoteHostIP & " JOIN :" & cName & vbCrLf

'send names list to client - we do this before the names is generated so own nick does not appear twice
names = Split(Channel(cIndex), ":")
names(0) = "" 'remove channel name from array
names(1) = Form1.sckTCP(Index).Tag 'remove channel modes and add own nick
names = Join(names, " ") 'rejoin it
names = Trim(names) 'trim spaces off of ends

'add to channel info the new nick
Channel(cIndex) = Channel(cIndex) & Form1.sckTCP(Index).Tag & ":"

'add to client data that he is in new channel
Client(Index) = Client(Index) & cName & ":"

'send names
Form1.sckTCP(Index).SendData ":" & ServerName & " 353 " & Form1.sckTCP(Index).Tag & " = " & cName & " :" & names & vbCrLf
'send end of names list
Form1.sckTCP(Index).SendData ":" & ServerName & " 366 " & Form1.sckTCP(Index).Tag & " " & cName & " :End of /NAMES list." & vbCrLf
End Function

Public Function PART(Index As Integer, The_Data)
tmpString = Split(The_Data, " ")
'remove channel from client info
Client(Index) = Replace(Client(Index), ":" & tmpString(1) & ":", ":")
'remove channel from channel info and send to clients that client left
For a = 1 To UBound(Channel)
 
 'increase a if channel(a) =""
 If Channel(a) = "" Then
  Do Until Channel(a) <> "" Or a = UBound(Channel)
   a = a + 1
  Loop
  If Channel(a) = "" Then Exit For 'fix for last item or channel in array = ""
 End If
 
 If Left(Channel(a), InStr(1, Channel(a), ":") - 1) = tmpString(1) Then
  'send to other clients that client left
  tmpString1 = Split(Channel(a), ":")
  For b = 1 To UBound(Client)
   For c = 2 To UBound(tmpString1) - 1
    If Form1.sckTCP(b).Tag = tmpString1(c) And Form1.sckTCP(b).State = 7 Then Form1.sckTCP(b).SendData ":" & Form1.sckTCP(Index).Tag & "!.@" & Form1.sckTCP(Index).RemoteHostIP & " PART " & tmpString(1) & vbCrLf
   Next c
  Next b
  'remove client from channel info
  Channel(a) = Replace(Channel(a), ":" & Form1.sckTCP(Index).Tag & ":", ":")
  'delete channel data if no one is int the channel
  tmpString = Split(Channel(a), ":")
  If UBound(tmpString) = 2 Then Channel(a) = ""
  Exit For
 End If
Next a
End Function

Public Function PRIVMSG(Index As Integer, The_Data)
'The_Data = privmsg #a :hi
'channel(whatever) = #channel:+modes:nick:nick:nick:nick:
'client(whatever) = nick:#channel:#channel:#channel:

tmpString = Split(The_Data, " ")

If Left(tmpString(1), 1) = "#" Then
'channel message
 For a = 1 To UBound(Client) - 1
  If InStr(1, Client(a), ":" & tmpString(1) & ":") <> 0 And Index <> a Then Form1.sckTCP(a).SendData ":" & Form1.sckTCP(Index).Tag & "!.@" & Form1.sckTCP(Index).RemoteHostIP & " " & The_Data & vbCrLf
 Next a
Else
'client message
 For a = 1 To UBound(Client) - 1
  If InStr(1, Client(a), tmpString(1) & ":") <> 0 Then Form1.sckTCP(a).SendData ":" & Form1.sckTCP(Index).Tag & "!.@" & Form1.sckTCP(Index).RemoteHostIP & " " & The_Data & vbCrLf
 Next a
End If
End Function

Public Function MODE(Index As Integer, The_Data)

tmpString = Split(The_Data, " ")

'join channel mode
If UBound(tmpString) = 1 Then
 For a = 1 To UBound(Channel)
 'increase a if channel(a) =""
 If Channel(a) = "" Then
  Do Until Channel(a) <> "" Or a = UBound(Channel)
   a = a + 1
  Loop
  If Channel(a) = "" Then Exit For 'fix for last item or channel in array = ""
 End If
  If Left(Channel(a), InStr(1, Channel(a), ":") - 1) = tmpString(1) Then
   tmpString1 = Split(Channel(a), ":")
   Form1.sckTCP(Index).SendData tmpString(1) & " " & tmpString1(1) & vbCrLf
   Exit For
  End If
 Next a
End If

   

End Function

Public Function Socket_Close(Index As Integer)
tmpString = Split(Client(Index), ":")
'make client part all channels first
For a = 1 To UBound(tmpString)
 Call PART(Index, "PART " & tmpString(a))
Next a
'clear the client data
Client(Index) = ""
'clear the socket tag
Form1.sckTCP(Index).Tag = ""
'close the socket
Form1.sckTCP(Index).Close
End Function

Public Function NOTICE(Index As Integer, The_Data)
'
End Function

