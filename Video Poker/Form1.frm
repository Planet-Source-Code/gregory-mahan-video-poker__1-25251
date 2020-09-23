VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video Poker"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6285
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   Begin VB.Frame Frame1 
      Caption         =   "Scoring"
      Height          =   1815
      Left            =   135
      TabIndex        =   13
      Top             =   3015
      Width           =   5955
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "5"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   18
         Left            =   3825
         TabIndex        =   32
         Top             =   1170
         Width           =   450
      End
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   17
         Left            =   3825
         TabIndex        =   31
         Top             =   900
         Width           =   450
      End
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "15"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   16
         Left            =   3825
         TabIndex        =   30
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "20"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   15
         Left            =   3825
         TabIndex        =   29
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "25"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   14
         Left            =   1395
         TabIndex        =   28
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "40"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   13
         Left            =   1395
         TabIndex        =   27
         Top             =   1170
         Width           =   450
      End
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "125"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   12
         Left            =   1395
         TabIndex        =   26
         Top             =   900
         Width           =   450
      End
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "250"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   11
         Left            =   1395
         TabIndex        =   25
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lblScoring 
         Alignment       =   1  'Right Justify
         Caption         =   "2000"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   10
         Left            =   1395
         TabIndex        =   24
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "Each turn costs $5 to play"
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   9
         Left            =   2790
         TabIndex        =   23
         Top             =   1440
         Width           =   1845
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "1 Pair"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   8
         Left            =   2790
         TabIndex        =   22
         Top             =   1170
         Width           =   405
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "2 Pair"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   2790
         TabIndex        =   21
         Top             =   900
         Width           =   405
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "3 of a Kind"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   2790
         TabIndex        =   20
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "Straight"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   2790
         TabIndex        =   19
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "Flush"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   18
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "Full House"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   17
         Top             =   1170
         Width           =   750
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "4 of a Kind"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   16
         Top             =   900
         Width           =   765
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "Straight Flush"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   15
         Top             =   630
         Width           =   960
      End
      Begin VB.Label lblScoring 
         AutoSize        =   -1  'True
         Caption         =   "Royal Flush"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   20
      Left            =   2880
      Top             =   1305
   End
   Begin VB.ListBox lstSuits 
      Height          =   1230
      ItemData        =   "Form1.frx":030A
      Left            =   1350
      List            =   "Form1.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstVals 
      Height          =   1230
      Left            =   495
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1665
      Tag             =   "0"
      Top             =   1305
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   495
      Tag             =   "0"
      Top             =   1305
   End
   Begin VB.PictureBox PicCard 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   4
      Left            =   5040
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   5
      Top             =   765
      Width           =   1065
   End
   Begin VB.PictureBox PicCard 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   3
      Left            =   3825
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   765
      Width           =   1065
   End
   Begin VB.PictureBox PicCard 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   2
      Left            =   2610
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   3
      Top             =   765
      Width           =   1065
   End
   Begin VB.PictureBox PicCard 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   1
      Left            =   1395
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   765
      Width           =   1065
   End
   Begin VB.PictureBox PicCard 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   0
      Left            =   225
      ScaleHeight     =   1440
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   765
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Draw!"
      Default         =   -1  'True
      Height          =   510
      Left            =   2565
      TabIndex        =   0
      Top             =   2385
      Width           =   1185
   End
   Begin VB.Label lblTurn 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1125
      TabIndex        =   12
      Top             =   2430
      Width           =   1230
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Turn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   225
      TabIndex        =   11
      Top             =   2430
      Width           =   825
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   225
      TabIndex        =   8
      Top             =   135
      Width           =   5865
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   4230
      TabIndex        =   7
      Top             =   2430
      Width           =   285
   End
   Begin VB.Label lblWinnings 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4590
      TabIndex        =   6
      Top             =   2430
      Width           =   1500
   End
   Begin VB.Shape CardSel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1485
      Index           =   4
      Left            =   5025
      Top             =   750
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Shape CardSel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1485
      Index           =   3
      Left            =   3810
      Top             =   750
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Shape CardSel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1485
      Index           =   2
      Left            =   2595
      Top             =   750
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Shape CardSel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1485
      Index           =   1
      Left            =   1380
      Top             =   750
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Shape CardSel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1485
      Index           =   0
      Left            =   210
      Top             =   750
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wCard As Long, hCard As Long, nCard As Integer

Private Sub Command1_Click()
Dim card As Integer
If Timer1.Enabled = True Then Exit Sub
If Timer2.Enabled = True Then Exit Sub
If Timer3.Enabled = True Then Exit Sub

Select Case Command1.Caption
Case "&Draw!"
    DisableAll
    Command1.Caption = "&Play Again"
    'get new values
    For card = 0 To 4
        If CardSel(card).Visible = True Then
            PicCard(card).Tag = Deck(nCard)
            nCard = nCard + 1
        End If
    Next card
    HideCards
    Do Until Timer2.Enabled = False
        DoEvents
    Loop
    ShowCards
    Do Until Timer1.Enabled = False
        DoEvents
    Loop
    CheckWinnings
    EnableAll
Case "&Play Again"
    DisableAll
    Command1.Caption = "&Draw!"
    ChangeScore -5
    lblTurn = lblTurn + 1
    lblMessage = ""
    Shuffle
    For card = 0 To 4
        CardSel(card).Visible = True
        CardSel(card).Tag = ""
        PicCard(card).Tag = Deck(card)
        cdtDraw PicCard(card).hdc, 0, 0, castle, CBACKS, 1
        PicCard(card).Refresh
    Next card
    nCard = 5 'next card from the deck.
    ShowCards
    EnableAll
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim chrVal As String
If Timer1.Enabled = True Then Exit Sub
If Timer2.Enabled = True Then Exit Sub
If Timer3.Enabled = True Then Exit Sub

chrVal = Chr(KeyAscii)
Select Case chrVal
Case "1", "2", "3", "4", "5"
    PicCard_Click (Val(chrVal) - 1)
End Select
End Sub

Private Sub Form_Load()
Me.Show
cdtInit wCard, hCard
NewGame
End Sub

Private Sub Form_Terminate()

cdtTerm
End Sub

Private Sub Form_Unload(Cancel As Integer)

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuNew_Click()
NewGame
End Sub

Private Sub PicCard_Click(Index As Integer)
CardSel(Index).Visible = (CardSel(Index).Visible = False)
If CardSel(Index).Visible = True Then
    CardSel(Index).Tag = "Flip"
Else
    CardSel(Index).Tag = ""
End If
End Sub

Private Sub PicCard_DblClick(Index As Integer)
CardSel(Index).Visible = (CardSel(Index).Visible = False)
If CardSel(Index).Visible = True Then
    CardSel(Index).Tag = "Flip"
Else
    CardSel(Index).Tag = ""
End If

End Sub

Sub ShowCards()
Timer1.Enabled = True
Do Until Timer1.Enabled = False
    DoEvents
Loop


End Sub
Sub HideCards()

Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim card As Integer
card = 0
For card = 0 To 4
    If CardSel(card).Visible = True Then
        'found one to turn over
        cdtDraw PicCard(card).hdc, 0, 0, Val(PicCard(card).Tag), CFACES, 1
        PicCard(card).Refresh
        CardSel(card).Visible = False
        Exit Sub
    End If
Next card
'all turned over!

Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
Dim card As Integer
card = 0
For card = 0 To 4
    If CardSel(card).Tag = "Flip" Then
        'found one to turn over
        cdtDraw PicCard(card).hdc, 0, 0, castle, CBACKS, 1
        PicCard(card).Refresh
        CardSel(card).Tag = ""
        Exit Sub
    End If
Next card
'all turned over!
Timer2.Enabled = False

End Sub
Sub CheckWinnings()
Dim i As Integer, vals(0 To 4) As Integer, suits(0 To 4) As Integer, bStraight As Boolean
lstVals.Clear
lstSuits.Clear
For i = 0 To 4
    Select Case PicCard(i).Tag
    Case cAce + cHearts, cAce + cDiamonds, cAce + cClubs, cAce + cSpades
        lstVals.AddItem "01"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cAce
    Case cTwo + cHearts, cTwo + cDiamonds, cTwo + cClubs, cTwo + cSpades
        lstVals.AddItem "02"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cTwo
    Case cThree + cHearts, cThree + cDiamonds, cThree + cClubs, cThree + cSpades
        lstVals.AddItem "03"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cThree
    Case cFour + cHearts, cFour + cDiamonds, cFour + cClubs, cFour + cSpades
        lstVals.AddItem "04"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cFour
    Case cFive + cHearts, cFive + cDiamonds, cFive + cClubs, cFive + cSpades
        lstVals.AddItem "05"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cFive
    Case cSix + cHearts, cSix + cDiamonds, cSix + cClubs, cSix + cSpades
        lstVals.AddItem "06"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cSix
    Case cSeven + cHearts, cSeven + cDiamonds, cSeven + cClubs, cSeven + cSpades
        lstVals.AddItem "07"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cSeven
    Case cEight + cHearts, cEight + cDiamonds, cEight + cClubs, cEight + cSpades
        lstVals.AddItem "08"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cEight
    Case cNine + cHearts, cNine + cDiamonds, cNine + cClubs, cNine + cSpades
        lstVals.AddItem "09"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cNine
    Case cTen + cHearts, cTen + cDiamonds, cTen + cClubs, cTen + cSpades
        lstVals.AddItem "10"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cTen
    Case cJack + cHearts, cJack + cDiamonds, cJack + cClubs, cJack + cSpades
        lstVals.AddItem "11"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cJack
    Case cQueen + cHearts, cQueen + cDiamonds, cQueen + cClubs, cQueen + cSpades
        lstVals.AddItem "12"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cQueen
    Case cKing + cHearts, cKing + cDiamonds, cKing + cClubs, cKing + cSpades
        lstVals.AddItem "13"
        'subtract the card value, leaving only the suit.
        lstSuits.AddItem PicCard(i).Tag - cKing
    End Select
Next i

For i = 0 To 4
    vals(i) = Val(lstVals.List(i))
    suits(i) = Val(lstSuits.List(i))
Next i
'now check for winners!

'royal flush 2000 'AKQJ10 all same suit
'also check for ace-high straight
If vals(0) = 1 And vals(1) = 10 And vals(2) = 11 And vals(3) = 12 And vals(4) = 13 Then
    If suits(0) = suits(4) Then
        lblMessage = "Royal Flush!!!"
        ChangeScore 2000
        Exit Sub
    Else
        lblMessage = "Straight!"
        ChangeScore 20
        Exit Sub
    End If
End If

'check for a straight. There are 2 kinds
'straight flush 250 straight with same suit
'straight 20 just a straight
'every card is 1 value more than the previous card.
'except in ace-high straights, which are taken care of
'in the royal flush check anwyay.
bStraight = True
For i = 0 To 3
    If vals(i + 1) - vals(i) <> 1 Then 'cards are more than 1 digit away from each other in value
        bStraight = False
        Exit For
    End If
Next i
'if bstraight is still true, the cards are in incremental order
If bStraight = True Then
    If suits(0) = suits(4) Then 'straight flush
        lblMessage = "Straight Flush!"
        ChangeScore 250
    Else 'normal straight
        lblMessage = "Straight!"
        ChangeScore 20
    End If
    Exit Sub
End If

'4 of a kind 125
'which can be xxxx*, or *xxxx
If (vals(0) = vals(3)) Or (vals(1) = vals(4)) Then  'middle values have to be the same value, because it's sorted
    'we have 3 of a kind
    lblMessage = "Four of a Kind!"
    ChangeScore 125
    Exit Sub
End If

'full house 40
'can be xxxyy or yyxxx only
If ((vals(0) = vals(2)) And (vals(3) = vals(4))) Or ((vals(0) = vals(1)) And (vals(2) = vals(4))) Then
    'we have a full house
    lblMessage = "Full House!"
    ChangeScore 40
    Exit Sub
End If

'flush 25
'all cards the same suit
If suits(0) = suits(4) Then
    'we have 3 of a kind
    lblMessage = "Flush!"
    ChangeScore 25
    Exit Sub
End If

'check for 3 of a kind 15
'which can be xxx**, *xxx*, or **xxx
If (vals(0) = vals(2)) Or (vals(1) = vals(3)) Or (vals(2) = vals(4)) Then 'middle value has to be the same value, because it's sorted
    'we have 3 of a kind
    lblMessage = "Three of a Kind!"
    ChangeScore 15
    Exit Sub
End If

'check for 2 pair 10
'which can be xx*yy *xxyy or xxyy*
If ((vals(0) = vals(1)) And (vals(3) = vals(4))) Or ((vals(1) = vals(2)) And (vals(3) = vals(4))) Or ((vals(0) = vals(1)) And (vals(2) = vals(3))) Then
    'we have 2 pair
    lblMessage = "Two Pair!"
    ChangeScore 10
    Exit Sub
End If

'check for a single pair 5
'a pair can be xx***, *xx**, **xx*, or ***xx
'but they will always be next to each other
If (vals(0) = vals(1)) Or (vals(1) = vals(2)) Or (vals(2) = vals(3)) Or (vals(3) = vals(4)) Then
    'we have a pair
    lblMessage = "One Pair!"
    ChangeScore 5
    Exit Sub
End If

lblMessage = "Bust!"
If lblWinnings = "0" Then
    lblMessage = "Game Over!"
    Command1.Enabled = False
End If
End Sub
Sub NewGame()
Dim i As Integer
DisableAll
lblMessage = ""
Shuffle
For i = 0 To 4
    CardSel(i).Visible = True
    CardSel(i).Tag = ""
    PicCard(i).Tag = Deck(i)
    cdtDraw PicCard(i).hdc, 0, 0, castle, CBACKS, 1
    PicCard(i).Refresh
Next i
nCard = 5 'next card from the deck.

lblWinnings = "50"
ChangeScore -5
ShowCards
Command1.Caption = "&Draw!"
Command1.Enabled = True
EnableAll
End Sub

Sub ChangeScore(intAmount As Integer)
Dim i As Integer, intIncr As Integer
If intAmount > 500 Then
    intIncr = 5
Else
    intIncr = 1
End If

If intAmount > 0 Then
    For i = 1 To intAmount
        lblWinnings = lblWinnings + intIncr
        lblWinnings.Refresh
        Timer3.Enabled = True
        Do Until Timer3.Enabled = False
            DoEvents
        Loop
    Next i
Else
    For i = 1 To Abs(intAmount)
        lblWinnings = lblWinnings - intIncr
        lblWinnings.Refresh
        Timer3.Enabled = True
        Do Until Timer3.Enabled = False
            DoEvents
        Loop
    Next i
End If

End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
End Sub
Sub DisableAll()
Dim control As control
For Each control In Form1.Controls
    If TypeName(control) <> "Shape" And TypeName(control) <> "Frame" And TypeName(control) <> "Label" And control.Name <> "mnuDash1" Then
        control.Enabled = False
    End If
Next control
End Sub
Sub EnableAll()
Dim control As control
For Each control In Form1.Controls
    If TypeName(control) <> "Shape" And TypeName(control) <> "Frame" And TypeName(control) <> "Label" And control.Name <> "mnuDash1" Then
        control.Enabled = True
    End If
Next control

End Sub
