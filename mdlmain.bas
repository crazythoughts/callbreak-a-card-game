Attribute VB_Name = "mdlmain"
Public player_name As String
Private counter As Integer
Public deck(52) As New cards
Public del_cards As Integer
Public winner_name As String
Public Sub Main()
    del_cards = 0
    Dim i As Integer
    i = 1
    For counter = 1 To 13 Step 1
        If i > 13 Then
            i = 1
        End If
        deck(counter).id = counter
        deck(counter).name = "spade"
        deck(counter).cvalue = i
        deck(counter).img = LoadPicture(App.Path & "\cards\spade\gif\" & i & ".gif")
        i = i + 1
    Next counter
    For counter = 14 To 26 Step 1
        If i > 13 Then
            i = 1
        End If
        deck(counter).id = counter
        deck(counter).name = "hearts"
        deck(counter).cvalue = i
        deck(counter).img = LoadPicture(App.Path & "\cards\hearts\gif\" & i & ".gif")
        i = i + 1
    Next counter
    For counter = 27 To 39 Step 1
        If i > 13 Then
            i = 1
        End If
        deck(counter).id = counter
        deck(counter).name = "clubs"
        deck(counter).cvalue = i
        deck(counter).img = LoadPicture(App.Path & "\cards\clubs\gif\" & i & ".gif")
        i = i + 1
    Next counter
    For counter = 40 To 52 Step 1
        If i > 13 Then
            i = 1
        End If
        deck(counter).id = counter
        deck(counter).name = "diamond"
        deck(counter).cvalue = i
        deck(counter).img = LoadPicture(App.Path & "\cards\diamond\gif\" & i & ".gif")
        i = i + 1
    Next counter
    frmSplash.Show
End Sub
