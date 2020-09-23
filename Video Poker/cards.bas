Attribute VB_Name = "Module1"
Option Explicit

'Delarations and utilities for using CARDS.DLL
'Actions for CdtDraw/Ext
' use in the nDraw field
Global Const CFACES = 0
Global Const CBACKS = 1
Global Const CINVERT = 2
'Card Numbers
' use in the nCard field
'from 0 to 51 [Ace (club,diamond,heart,spade), Deuce, ... , King]
'Card Backs
' use in the nCard field
' CAUTION: when nCard > 53 then nDraw must be = 1 (C_BACKS)
Global Const crosshatch = 53
Global Const weave1 = 54
Global Const weave2 = 55
Global Const robot = 56
Global Const flowers = 57
Global Const vine1 = 58
Global Const vine2 = 59
Global Const fish1 = 60
Global Const fish2 = 61
Global Const shells = 62
Global Const castle = 63
Global Const island = 64
Global Const cardhand = 65
Global Const UNUSED = 66
Global Const THE_X = 67
Global Const THE_O = 68

'constants for the suits use suit+value to get a card
Global Const cClubs = 0
Global Const cDiamonds = 1
Global Const cHearts = 2
Global Const cSpades = 3

'constants for the value
Global Const cAce = 0
Global Const cTwo = 4
Global Const cThree = 8
Global Const cFour = 12
Global Const cFive = 16
Global Const cSix = 20
Global Const cSeven = 24
Global Const cEight = 28
Global Const cNine = 32
Global Const cTen = 36
Global Const cJack = 40
Global Const cQueen = 44
Global Const cKing = 48

'Initialization
' call before anything else. Returns the default
' width and height for the cards, in pixels.
Declare Function cdtInit Lib "cards.dll" (dx As Long, dy As Long) As Long
'CdtDraw used to draw a card with the default size
'at a specified location in a form, picture box or whatever.
'It can draw any of the 52 faces an 13 different Back designs,
'as well as pile markers such as the X and O. Cards can also
'be drawn in the negative image, eg to show selection.
'xOrg = x origin in pixels
'yOrg = y origin in pixels
'nCard = one of the Card Back constants or a card number 0 to 51
'nDraw = one of the Action constants
'nColor = The highlight color
Declare Function cdtDraw Lib "cards.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal iCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
'CdtDrawExt used to draw a card in any size
'Much the same as CdtDraw, but you can specify the height & width
'of the card, as well as location.
'nWidth = Width of card in pixels
'nHeight = Height of card in pixels.
Declare Function cdtDrawExt Lib "cards.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal ordCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
'CdtTerm should be called when the program terminates.
' Primarily it releases memory back to Windows.
Declare Function cdtTerm Lib "cards.dll" () As Long
'cdtAnimate animates the backs of cards by overlaying part of
'the card back with an alternative bitmap. It creates
'effects: blinking lights on the robot, the sun donning
'sunglasses, bats flying across the castle, and a card sliding
'out of a sleeve. The function works only for cards of normal
'size drawn with cdtDraw. To draw each state, start with
'iState set to 0 and increment through until cdtAnimate
'returns 0.
Declare Function cdtAnimate Lib "cards.dll" (ByVal hdc As Long, ordCard As Long, ByVal X As Long, ByVal Y As Long, spr As Long) As Long

Global Deck(0 To 51) As Integer

Sub Shuffle()
Dim intCount As Integer, Half1() As Integer, Half2() As Integer
Dim cutPoint As Integer, h1Count As Integer, h2Count As Integer, sCount As Integer
'sure, I could make a quick and easy shuffle with a 3 line rnd() loop, but I prefer to model it more real-world
Randomize
'order deck
For intCount = 0 To 51
    Deck(intCount) = intCount
Next intCount

For sCount = 1 To 7
    'cut deck kinda in half
    cutPoint = 26
    cutPoint = cutPoint + (Fix(Rnd * 15) - 7) 'deviate +-7 from our cutpoint
    
    If cutPoint > 25 Then 'half1 is always larger half of deck
        ReDim Half1(cutPoint)
        ReDim Half2(51 - cutPoint)
    Else
        ReDim Half1(51 - cutPoint)
        ReDim Half2(cutPoint)
    End If
    'split values into the two deckhalves
    For intCount = 0 To 51
        If intCount <= UBound(Half1) Then
            Half1(intCount) = Deck(intCount)
        Else
            Half2(intCount - (UBound(Half1) + 1)) = Deck(intCount)
        End If
    Next intCount
    'now interleave  the two halfs together into the singular deck
    For intCount = 0 To 51
        Deck(intCount) = -1
    Next intCount
    intCount = 0
    h1Count = 0
    h2Count = 0
    Do
        '80% chance of dropping card from half1 into the deck
        If Fix(Rnd * 100) <= 79 Then
            If h1Count <= UBound(Half1) Then 'but only if there are still cards left in this half
                Deck(intCount) = Half1(h1Count)
                'If Deck(intCount) = 0 Then Stop
                h1Count = h1Count + 1
                intCount = intCount + 1
            End If
        End If
        
        '80% chance of dropping card from half2 into the deck
        If Fix(Rnd * 100) <= 79 Then
            If h2Count < UBound(Half2) Then 'but only if there are still cards left in this half
                Deck(intCount) = Half2(h2Count)
                h2Count = h2Count + 1
                intCount = intCount + 1
            End If
        End If
    Loop Until intCount = 52
    
    'cut the deck
    'go thru same splitting routine as when we interleave, only this time split the deck anywhere leaving at least 2 cards in the smaller stack
    cutPoint = Fix(Rnd * 51) + 1
    ReDim Half1(cutPoint)
    ReDim Half2(51 - cutPoint)
    'split values into the two deckhalves
    For intCount = 0 To 51
        If intCount <= UBound(Half1) Then
            Half1(intCount) = Deck(intCount)
        Else
            Half2(intCount - (UBound(Half1) + 1)) = Deck(intCount)
        End If
    Next intCount
    'recombine the two values, putting half2 on top
    For intCount = 0 To 51
        If intCount < UBound(Half2) Then
            Deck(intCount) = Half2(intCount)
        Else
            Deck(intCount) = Half1(intCount - (UBound(Half2)))
        End If
    Next intCount
Next sCount
End Sub

