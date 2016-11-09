VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   Caption         =   "Server"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   1860
   LinkTopic       =   "Form2"
   ScaleHeight     =   1470
   ScaleWidth      =   1860
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrgame 
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin MSWinsockLib.Winsock tcpgame 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private con_request As Integer
Private shuffled(52) As New cards
Private data_recieve As String
Private Players(4) As New user
Private Game_starter As Integer
Private round_turn As Integer
Private comp_cards(4) As New comp_class
Private highest_card As New comp_class
Public sendingcomplete As Boolean
Public turn_counter As Integer
Public round_counter As Integer
Public game_counter As Integer
Public highest_score As Double
Public winner_id As Integer
Public Sub waitforwinsock()
While Not sendingcomplete
DoEvents
Wend
End Sub
Private Sub Form_Load()
    'for the game server winsock
        tcpgame(0).Close
        tcpgame(0).LocalPort = 6001
        tcpgame(0).Listen
        Game_starter = 1
        round_turn = 1
        turn_counter = 0
        round_counter = 0
        game_counter = 0
End Sub
Private Sub tcpgame_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim counter As Integer
    Dim player_counter As Integer
    If Index = 0 Then
        con_request = con_request + 1
        Load tcpgame(con_request)
        tcpgame(con_request).LocalPort = 0
        tcpgame(con_request).Accept requestID
        'assigning the players
            Players(con_request).id = con_request
            Players(con_request).ccall = 0
            Players(con_request).win = 0
            Players(con_request).cpoints = 0
            Players(con_request).tpoints = 0
    End If
End Sub

'<=================================================================DATA ARRIVAL===========================================================================>
Private Sub tcpgame_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim totalmsg() As String
    Dim counter As Integer
    tcpgame(Index).GetData data_recieve
    totalmsg() = Split(data_recieve, "|")
    If totalmsg(0) = "message" Then
    sendingcomplete = False
        For counter = 1 To con_request Step 1
            If Index = 1 Then
                tcpgame(counter).SendData "message|" & Players(1).player_name & "|" & totalmsg(1)
            ElseIf Index = 2 Then
                tcpgame(counter).SendData "message|" & Players(2).player_name & "|" & totalmsg(1)
            ElseIf Index = 3 Then
                tcpgame(counter).SendData "message|" & Players(3).player_name & "|" & totalmsg(1)
            ElseIf Index = 4 Then
                tcpgame(counter).SendData "message|" & Players(4).player_name & "|" & totalmsg(1)
            End If
        Next counter
    waitforwinsock
    ElseIf totalmsg(0) = "call" Then
        Players(Index).ccall = CInt(totalmsg(1))
    sendingcomplete = False
        For counter = 1 To con_request Step 1
            tcpgame(counter).SendData "call|" & Index & "|" & Players(Index).ccall
        Next counter
    waitforwinsock
    ElseIf totalmsg(0) = "name" Then
        If Index = 1 Then
            Players(Index).player_name = totalmsg(1)
        ElseIf Index = 2 Then
            Players(Index).player_name = totalmsg(1)
        ElseIf Index = 3 Then
            Players(Index).player_name = totalmsg(1)
        ElseIf Index = 4 Then
            Players(Index).player_name = totalmsg(1)
        End If
    ElseIf totalmsg(0) = "cardthrown" Then
        turn_counter = turn_counter + 1
        comp_cards(turn_counter).id = CInt(totalmsg(1))
        comp_cards(turn_counter).holder = Index
        sendingcomplete = False
        For counter = 1 To con_request Step 1
            If counter = Index Then
                'Do nothing
            Else
                tcpgame(counter).SendData "cardthrown|" & Index & "|" & totalmsg(1)
            End If
        Next counter
        waitforwinsock
        round_turn = round_turn + 1
        If round_turn > 4 Then
            round_turn = 1
        End If
        If turn_counter = 4 Then
            Call Card_comparison
        Else
            sendingcomplete = False
            For counter = 1 To con_request Step 1
                If counter = round_turn Then
                    tcpgame(counter).SendData "RS|" & "ok"
                Else
                    tcpgame(counter).SendData "RS|" & "not"
                End If
            Next counter
            waitforwinsock
        End If
    End If
End Sub

'<===============================================================FUNCTION BLOCKS====================================================================>
Function Card_comparison()
    Dim counter_first As Integer
    Dim counter_second As Integer
    For counter_first = 1 To 4 Step 1
        For counter_second = 1 To 52 Step 1
            If deck(counter_second).id = comp_cards(counter_first).id Then
                comp_cards(counter_first).name = deck(counter_second).name
                comp_cards(counter_first).cval = deck(counter_second).cvalue
                Exit For
            End If
        Next counter_second
    Next counter_first
    highest_card.id = comp_cards(1).id
    highest_card.holder = comp_cards(1).holder
    highest_card.name = comp_cards(1).name
    highest_card.cval = comp_cards(1).cval
    For counter_first = 1 To 4 Step 1
        If (comp_cards(counter_first).name = highest_card.name Or comp_cards(counter_first).name = "spade") Then
            If comp_cards(counter_first).name = "spade" And comp_cards(counter).name = highest_card.name Then
                If comp_cards(counter_first).cval > highest_card.cval Then
                    highest_card.id = comp_cards(counter_first).id
                    highest_card.holder = comp_cards(counter_first).holder
                    highest_card.name = comp_cards(counter_first).name
                    highest_card.cval = comp_cards(counter_first).cval
                End If
            ElseIf comp_cards(counter_first).name = "spade" And comp_cards(counter_first).name <> highest_card.name Then
                highest_card.id = comp_cards(counter_first).id
                highest_card.holder = comp_cards(counter_first).holder
                highest_card.name = comp_cards(counter_first).name
                highest_card.cval = comp_cards(counter_first).cval
            Else
                If comp_cards(counter_first).name = highest_card.name Then
                    If comp_cards(counter_first).cval > highest_card.cval Then
                        highest_card.id = comp_cards(counter_first).id
                        highest_card.holder = comp_cards(counter_first).holder
                        highest_card.name = comp_cards(counter_first).name
                        highest_card.cval = comp_cards(counter_first).cval
                    End If
                End If
            End If
        End If
    Next counter_first
    turn_counter = 0
    For counter_second = 1 To 4 Step 1
        If Players(counter_second).id = highest_card.holder Then
            Players(counter_second).win = (Players(counter_second).win + 1)
            Exit For
        End If
    Next counter_second
    round_turn = highest_card.holder
    sendingcomplete = False
    For counter_first = 1 To con_request Step 1
        If counter_first = highest_card.holder Then
            tcpgame(counter_first).SendData "oho|" & highest_card.holder & "|" & Players(counter_second).win & "|" & comp_cards(1).id & "|" & comp_cards(2).id & "|" & comp_cards(3).id & "|" & comp_cards(4).id
        Else
            tcpgame(counter_first).SendData "oho|" & highest_card.holder & "|" & Players(counter_second).win
        End If
    Next counter_first
    waitforwinsock
    sendingcomplete = False
    For counter = 1 To con_request Step 1
        If counter = round_turn Then
            tcpgame(counter).SendData "RS|" & "ok"
        Else
            tcpgame(counter).SendData "RS|" & "not"
        End If
    Next counter
    waitforwinsock
    del_cards = del_cards + 1
    round_counter = round_counter + 1
    If round_counter = 13 Then
        Call Calc_points
    End If
End Function
Private Function Calc_points()
    Dim counter As Integer
    Dim counter2 As Integer
    Dim extra As Double
    For counter = 1 To 4 Step 1
        If Players(counter).win < Players(counter).ccall Then
            Players(counter).cpoints = 0
            Players(counter).tpoints = Players(counter).tpoints + Players(counter).cpoints
        ElseIf Players(counter).win = Players(counter).ccall Then
            Players(counter).cpoints = Players(counter).ccall
            Players(counter).tpoints = Players(counter).tpoints + Players(counter).cpoints
        Else
            extra = Players(counter).win - Players(counter).ccall
            Players(counter).cpoints = Players(counter).ccall + (extra / 10)
            Players(counter).tpoints = Players(counter).tpoints + Players(counter).cpoints
        End If
    Next counter
    game_counter = game_counter + 1
    sendingcomplete = False
    For counter = 1 To con_request Step 1
        For counter2 = 1 To 4 Step 1
            If counter2 = 1 Then
                tcpgame(counter).SendData "eor|" & game_counter & "|" & Players(counter2).cpoints & "|" & Players(counter2).tpoints & "|"
            Else
                tcpgame(counter).SendData Players(counter2).cpoints & "|" & Players(counter2).tpoints & "|"
            End If
        Next counter2
    Next counter
    waitforwinsock
    round_counter = 0
    del_cards = 0
    For counter = 1 To 4
        Players(counter).win = 0
    Next counter
    'game_counter = game_counter + 1
    Game_starter = Game_starter + 1
    If Game_starter > 4 Then
        Game_starter = 1
    End If
    If game_counter < 5 Then
        tmrgame.Enabled = True
    Else
        highest_score = Players(1).tpoints
        For counter = 1 To 4 Step 1
            If Players(counter).tpoints > highest_score Then
                highest_score = Players(counter).tpoints
                winner_id = Players(counter).id
            End If
        Next counter
    sendingcomplete = False
        For counter = 1 To 4 Step 1
            tcpgame(counter).SendData "WG|" & winner_id
        Next counter
        waitforwinsock
    End If
End Function
Function shuffle_cards()
    Dim col As Collection
    Dim X As Integer
    Dim counter As Integer
    Randomize
    Set col = New Collection
    For counter = 1 To 52
        col.Add counter
    Next counter
    For counter = 1 To 52
        X = RandomInteger(1, col.Count)
        shuffled(counter).name = deck(col.Item(X)).name
        shuffled(counter).id = deck(col.Item(X)).id
        shuffled(counter).cvalue = deck(col.Item(X)).cvalue
        shuffled(counter).img = deck(col.Item(X)).img
        col.Remove (X)
    Next counter
End Function
Private Function RandomInteger(Lowerbound As Integer, Upperbound As Integer) As Integer
RandomInteger = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function
Private Function dis_cards()
    Dim counter_cards As Integer
    Dim counter As Integer
    Dim counter_player As Integer
    For counter_player = 1 To 4 Step 1
        sendingcomplete = False
        If counter_player = 1 Then
            If tcpgame(counter_player).State = sckConnected Then
                For counter_cards = 1 To 13 Step 1
                    If counter_cards = 1 Then
                        tcpgame(counter_player).SendData "distrib|" & shuffled(counter_cards).id & "|"
                    Else
                        tcpgame(counter_player).SendData shuffled(counter_cards).id & "|"
                    End If
                Next counter_cards
            End If
        ElseIf counter_player = 2 Then
            If tcpgame(counter_player).State = sckConnected Then
                For counter_cards = 14 To 26 Step 1
                    If counter_cards = 14 Then
                        tcpgame(counter_player).SendData "distrib|" & shuffled(counter_cards).id & "|"
                    Else
                        tcpgame(counter_player).SendData shuffled(counter_cards).id & "|"
                    End If
                Next counter_cards
            End If
        ElseIf counter_player = 3 Then
            If tcpgame(counter_player).State = sckConnected Then
                For counter_cards = 27 To 39 Step 1
                    If counter_cards = 27 Then
                        tcpgame(counter_player).SendData "distrib|" & shuffled(counter_cards).id & "|"
                    Else
                        tcpgame(counter_player).SendData shuffled(counter_cards).id & "|"
                    End If
                Next counter_cards
            End If
        Else
            If tcpgame(counter_player).State = sckConnected Then
                For counter_cards = 40 To 52 Step 1
                    If counter_cards = 40 Then
                        tcpgame(counter_player).SendData "distrib|" & shuffled(counter_cards).id & "|"
                    Else
                        tcpgame(counter_player).SendData shuffled(counter_cards).id & "|"
                    End If
                Next counter_cards
            End If
        End If
    Next counter_player
    waitforwinsock
    round_turn = Game_starter
    sendingcomplete = False
    For counter = 1 To con_request Step 1
        If counter = Game_starter Then
            tcpgame(counter).SendData "GS|" & "ok"
        Else
            tcpgame(counter).SendData "GS|" & "not"
        End If
    Next counter
    waitforwinsock
End Function
Private Sub tcpgame_SendComplete(Index As Integer)
sendingcomplete = True
End Sub
Private Sub tmrgame_Timer()
    If con_request = 4 Then
        sendingcomplete = False
        For counter = 1 To 4 Step 1
            For player_counter = 1 To 4 Step 1
                If player_counter = 1 Then
                'MsgBox Players(1).win
                    tcpgame(counter).SendData "playerinfo|" & Players(player_counter).id & "|" & Players(player_counter).player_name & "|" & Players(player_counter).ccall & "|" & Players(player_counter).tpoints & "|" & Players(player_counter).cpoints & "|" & Players(player_counter).win & "|"
                Else
                    tcpgame(counter).SendData Players(player_counter).id & "|" & Players(player_counter).player_name & "|" & Players(player_counter).ccall & "|" & Players(player_counter).tpoints & "|" & Players(player_counter).cpoints & "|" & Players(player_counter).win & "|"
                End If
            Next player_counter
        Next counter
        waitforwinsock
        sendingcomplete = False
        Call shuffle_cards
        Call dis_cards
        waitforwinsock
        tmrgame.Enabled = False
    End If
End Sub
