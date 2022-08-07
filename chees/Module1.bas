Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public FoundWindows As String, ZSend As String
Public ZTop As Integer, ZLeft As Integer, ZWidth As Integer, ZHeight As Integer
Public AQZ_H As Integer, AQZ_M As Integer, AQZ_S As Integer, ZTempPath As String

' Object reference to the DLL that contains the resources to be loaded.
Public clsCTARS As Object
Public Const Indicator = ":':"

'UDP Port
Public Client As New Collection
Public Names As New Collection
Public RmIP As String, RmPt As String

'Sound
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'  flag values for uFlags parameter
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Public Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START = 0  '  must be > 4096 to keep strings in same section of resource file
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside
Public Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Public Const SND_TYPE_MASK = &H170007

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowLW Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Const GWL_ID = (-12)
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

'My Defs
Public ZWhosTurn As Integer, ZRotated As Boolean, ZToggle As Integer, BoardFrom As Integer, BoardTo As Integer, BlackInCheck As Boolean, WhiteInCheck As Boolean, BlackCheckMate As Boolean, WhiteCheckMate As Boolean
Public LegalMove As Boolean, OpInProgress As Boolean, OpInProgress2 As Boolean, WhatPiece As String, BlackKingHasMoved As Boolean, WhiteKingHasMoved As Boolean
Public WhiteCanCastleKingSide As Boolean, WhiteCanCastleQueenSide As Boolean, BlackCanCastleKingSide As Boolean, BlackCanCastleQueenSide As Boolean, TriedToCastle As Boolean
Public ZBoardFrom As Integer, ZBoardTo As Integer, ZBoardFromTo As String, ZLastMove As String, ZGameInProcess As Boolean, ZMovingPiece As Boolean
Public ZGuestMode As Boolean, ZLastMessage As String
Public Const CWhite = 1
Public Const CBlack = 2


Public Sub SendOutIP()
On Local Error Resume Next
Err.Clear

'Send a text message to all clients in collection/listbox
Dim X As Integer

'Loop through all IP in listbox and get the right Users IP
For X = 0 To RJSoftChess.lName.ListCount - 1
    'Select each IP
    RJSoftChess.lName.ListIndex = X
    'Set IP and Port to send to
    RmIP = RJSoftChess.lName.Text
    RmPt = Client(RmIP)
    RJSoftChess.Wsck.RemoteHost = RmIP
    RJSoftChess.Wsck.RemotePort = RmPt
    'Send text message
    RJSoftChess.Wsck.SendData ZSend
Next
Err.Clear

End Sub

Public Sub FindPath()
On Local Error Resume Next
Err.Clear

Dim ZTempPath As String, X As Integer

ZTempPath = String(145, 0)
X = GetWindowsDirectory(ZTempPath, 145)
ZTempPath = Left(ZTempPath, X)

If Right(ZTempPath, 1) <> "\" Then
    FoundWindows = ZTempPath + "\"
Else
    FoundWindows = ZTempPath
End If

End Sub
Public Function CheckForLegalMove(XMove As Boolean)
On Local Error Resume Next
Err.Clear

Dim ZFrom As Integer, ZTo As Integer, ZFrom2 As Integer, ZTo2 As Integer, XY As Integer, Y As Integer, ZDirection As Long, ZTempKing As String

LegalMove = False
TriedToCastle = False

'Pawn
If WhatPiece = "PN" Then
    'check 2
    If (Abs(Int(BoardFrom / 8)) = 1) Or (Abs(Int(BoardFrom / 8)) = 6) Then
        If (Abs(BoardFrom - (Int(BoardFrom / 8) * 8) = Abs(BoardTo - (Int(BoardTo / 8) * 8)))) And (Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) = 2) And (RJSoftChess.Board(BoardTo).Tag = "") Then LegalMove = True
        If Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) > 2 Then Exit Function
     Else
        If Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) > 1 Then Exit Function
    End If
    'check 1
    If ZRotated = True Then
        If ZWhosTurn = CWhite Then
            If BoardTo < BoardFrom Then Exit Function
         Else
            If BoardFrom < BoardTo Then Exit Function
        End If
        If ZWhosTurn = CWhite Then
            If Abs(BoardFrom - (Int(BoardFrom / 8) * 8) = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Int(BoardTo / 8) - Int(BoardFrom / 8)) = 1 Then LegalMove = True
            If BoardTo >= 56 And BoardTo <= 63 And XMove = True Then
                RJSoftChess.Board(BoardFrom).Picture = RJSoftChess.Master_White(1).Picture
                RJSoftChess.Board(BoardFrom).Tag = "WQU"
            End If
         Else
            If Abs(BoardFrom - (Int(BoardFrom / 8) * 8) = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Int(BoardTo / 8) - Int(BoardFrom / 8)) = -1 Then LegalMove = True
            If BoardTo >= 0 And BoardTo <= 7 And XMove = True Then
                RJSoftChess.Board(BoardFrom).Picture = RJSoftChess.Master_Black(1).Picture
                RJSoftChess.Board(BoardFrom).Tag = "BQU"
            End If
        End If
     Else
        If ZWhosTurn = CWhite Then
            If BoardFrom < BoardTo Then Exit Function
         Else
            If BoardTo < BoardFrom Then Exit Function
        End If
        If ZWhosTurn = CWhite Then
            If Abs(BoardFrom - (Int(BoardFrom / 8) * 8) = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Int(BoardTo / 8) - Int(BoardFrom / 8)) = -1 Then LegalMove = True
            If BoardTo >= 0 And BoardTo <= 7 And XMove = True Then
                RJSoftChess.Board(BoardFrom).Picture = RJSoftChess.Master_White(1).Picture
                RJSoftChess.Board(BoardFrom).Tag = "WQU"
            End If
         Else
            If Abs(BoardFrom - (Int(BoardFrom / 8) * 8) = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Int(BoardTo / 8) - Int(BoardFrom / 8)) = 1 Then LegalMove = True
            If BoardTo >= 56 And BoardTo <= 63 And XMove = True Then
                RJSoftChess.Board(BoardFrom).Picture = RJSoftChess.Master_Black(1).Picture
                RJSoftChess.Board(BoardFrom).Tag = "BQU"
            End If
        End If
    End If
    'Can take the piece?
    If RJSoftChess.Board(BoardTo).Tag <> "" Then
        'Kill at angle
        If (Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) + 1 = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) > 0) Then LegalMove = True
        If (Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) - 1 = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) > 0) Then LegalMove = True
        'Can't kill head on.  Must be at an angle
        If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) - Abs(BoardTo - (Int(BoardTo / 8) * 8)) = 0 Then LegalMove = False
        If LegalMove = False Then
            'Must not be backwards
            If ZRotated = True Then
                If ZWhosTurn = CWhite Then
                    ZDirection = ((BoardFrom - BoardTo) / Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)))
                    If ZDirection = 7 Or ZDirection = 9 Then LegalMove = False
                 Else
                    ZDirection = ((BoardFrom - BoardTo) / Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)))
                    If ZDirection = -7 Or ZDirection = -9 Then LegalMove = False
                End If
             Else
                If ZWhosTurn = CWhite Then
                    ZDirection = ((BoardFrom - BoardTo) / Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)))
                    If ZDirection = -7 Or ZDirection = -9 Then LegalMove = False
                 Else
                    ZDirection = ((BoardFrom - BoardTo) / Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)))
                    If ZDirection = 7 Or ZDirection = 9 Then LegalMove = False
                End If
            End If
        End If
     Else
        If Abs(BoardFrom - (Int(BoardFrom / 8) * 8) = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Abs(Int(BoardTo / 8) - Int(BoardFrom / 8))) = 2 Then
            'Can't move backwards
            If ZRotated = True Then
                If ZWhosTurn = CWhite Then
                    If RJSoftChess.Board(BoardTo - 8).Tag <> "" Then LegalMove = False
                 Else
                    If RJSoftChess.Board(BoardTo + 8).Tag <> "" Then LegalMove = False
                End If
             Else
                If ZWhosTurn = CWhite Then
                    If RJSoftChess.Board(BoardTo + 8).Tag <> "" Then LegalMove = False
                 Else
                    If RJSoftChess.Board(BoardTo - 8).Tag <> "" Then LegalMove = False
                End If
            End If
        End If
    End If
    GoTo TestMove
    Exit Function
End If

'Rook
If WhatPiece = "RK" Then
    If (Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) > 0) Then
        LegalMove = True
        ZDirection = ((BoardFrom - BoardTo) / Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)))
    End If
    If (Abs(BoardFrom - (Int(BoardFrom * 8) / 8)) = Abs(BoardTo - (Int(BoardTo * 8) / 8))) And (Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) = 0) Then
        LegalMove = True
        ZDirection = BoardFrom - BoardTo
    End If
    If LegalMove = True Then
        GoSub TestUpDownSideSide
        GoTo TestMove
    End If
    Exit Function
End If

'Knight
If WhatPiece = "KN" Then
    ZFrom = BoardFrom - (Int(BoardFrom / 8) * 8)
    ZTo = BoardTo - (Int(BoardTo / 8) * 8)
    ZFrom2 = Int(BoardFrom / 8)
    ZTo2 = Int(BoardTo / 8)
    'check 1 right 2 down
    If ZTo = ZFrom + 1 And ZTo2 = ZFrom2 + 2 Then LegalMove = True
    'check 2 right 1 down
    If ZTo = ZFrom + 2 And ZTo2 = ZFrom2 + 1 Then LegalMove = True
    'check 1 right 2 up
    If ZTo = ZFrom + 1 And ZTo2 = ZFrom2 - 2 Then LegalMove = True
    'check 2 right 1 up
    If ZTo = ZFrom + 2 And ZTo2 = ZFrom2 - 1 Then LegalMove = True
    'check 1 left 2 down
    If ZTo = ZFrom - 1 And ZTo2 = ZFrom2 + 2 Then LegalMove = True
    'check 1 left 2 up
    If ZTo = ZFrom - 1 And ZTo2 = ZFrom2 - 2 Then LegalMove = True
    'check 2 left 1 down
    If ZTo = ZFrom - 2 And ZTo2 = ZFrom2 + 1 Then LegalMove = True
    'check 2 left 1 up
    If ZTo = ZFrom - 2 And ZTo2 = ZFrom2 - 1 Then LegalMove = True
    GoTo TestMove
    Exit Function
End If

'Bishop
If WhatPiece = "BP" Then
    If Abs(Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) - Abs(BoardTo - (Int(BoardTo / 8) * 8))) = Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) Then
        LegalMove = True
        GoSub TestAngle
        GoTo TestMove
    End If
    Exit Function
End If

'Queen
If WhatPiece = "QU" Then
    'Angle
    If Abs(Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) - Abs(BoardTo - (Int(BoardTo / 8) * 8))) = Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) Then
        LegalMove = True
        GoSub TestAngle
    End If
    'Up/Down/Side/Side
    If LegalMove = False Then
        If (Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = Abs(BoardTo - (Int(BoardTo / 8) * 8))) And (Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) > 0) Then
            LegalMove = True
            ZDirection = ((BoardFrom - BoardTo) / Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)))
        End If
        If (Abs(BoardFrom - (Int(BoardFrom * 8) / 8)) = Abs(BoardTo - (Int(BoardTo * 8) / 8))) And (Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)) = 0) Then
            LegalMove = True
            ZDirection = BoardFrom - BoardTo
        End If
        If LegalMove = True Then GoSub TestUpDownSideSide
    End If
    GoTo TestMove
    Exit Function
End If

'King
If WhatPiece = "KK" Then
    'Up/Down
    If BoardTo = BoardFrom + 8 Then LegalMove = True
    If BoardTo = BoardFrom - 8 Then LegalMove = True
    'Side to Side
    If BoardTo = BoardFrom + 1 Then LegalMove = True
    If BoardTo = BoardFrom - 1 Then LegalMove = True
    'Angle
    If BoardTo = BoardFrom + 7 Then LegalMove = True
    If BoardTo = BoardFrom - 7 Then LegalMove = True
    If BoardTo = BoardFrom + 9 Then LegalMove = True
    If BoardTo = BoardFrom - 9 Then LegalMove = True
    
    'Check to see if moving next to other King
    If ZWhosTurn = CWhite Then
        ZTempKing = "BKK"
     Else
        ZTempKing = "WKK"
    End If
    Err.Clear
    'Up/Down
    If RJSoftChess.Board(BoardTo + 8).Tag = ZTempKing Then
        If Err = 0 Then
            LegalMove = False
            Exit Function
        End If
    End If
    Err.Clear
    If RJSoftChess.Board(BoardTo - 8).Tag = ZTempKing Then
        If Err = 0 Then
            LegalMove = False
            Exit Function
        End If
    End If
    Err.Clear
    'Side to Side
    If RJSoftChess.Board(BoardTo + 1).Tag = ZTempKing Then
        If Err = 0 Then
            LegalMove = False
            Exit Function
        End If
    End If
    Err.Clear
    If RJSoftChess.Board(BoardTo - 1).Tag = ZTempKing Then
        If Err = 0 Then
            LegalMove = False
            Exit Function
        End If
    End If
    Err.Clear
    'Angle
    If RJSoftChess.Board(BoardTo + 7).Tag = ZTempKing Then
        If Err = 0 Then
            LegalMove = False
            Exit Function
        End If
    End If
    Err.Clear
    If RJSoftChess.Board(BoardTo - 7).Tag = ZTempKing Then
        If Err = 0 Then
            LegalMove = False
            Exit Function
        End If
    End If
    Err.Clear
    If RJSoftChess.Board(BoardTo + 9).Tag = ZTempKing Then
        If Err = 0 Then
            LegalMove = False
            Exit Function
        End If
    End If
    Err.Clear
    If RJSoftChess.Board(BoardTo - 9).Tag = ZTempKing Then
        If Err = 0 Then
            LegalMove = False
            Exit Function
        End If
    End If
    
    If ZWhosTurn = CWhite Then
        If ZRotated = True Then
            'Tried To Castle
            If (BoardFrom = 4 And BoardTo = 6) Or (BoardFrom = 4 And BoardTo = 2) Then TriedToCastle = True
            If WhiteKingHasMoved = False And WhiteInCheck = False Then
                'Castle?
                If BoardFrom = 4 And BoardTo = 6 Then
                    LegalMove = True
                    If XMove = True Then
                        'Castle King Side
                        RJSoftChess.Board(5).Picture = RJSoftChess.Board(7).Picture
                        RJSoftChess.Board(5).Tag = RJSoftChess.Board(7).Tag
                        RJSoftChess.Board(7).Picture = RJSoftChess.Master_Blank.Picture
                        RJSoftChess.Board(7).Tag = ""
                        WhiteCanCastleKingSide = False
                        WhiteCanCastleQueenSide = False
                    End If
                End If
                If BoardFrom = 4 And BoardTo = 2 Then
                    LegalMove = True
                    If XMove = True Then
                        'Castle Queen Side
                        RJSoftChess.Board(3).Picture = RJSoftChess.Board(0).Picture
                        RJSoftChess.Board(3).Tag = RJSoftChess.Board(0).Tag
                        RJSoftChess.Board(0).Picture = RJSoftChess.Master_Blank.Picture
                        RJSoftChess.Board(0).Tag = ""
                        WhiteCanCastleKingSide = False
                        WhiteCanCastleQueenSide = False
                    End If
                End If
            End If
            If LegalMove = True And XMove = True Then WhiteKingHasMoved = True
         Else
            'Tried To Castle
            If (BoardFrom = 59 And BoardTo = 57) Or (BoardFrom = 59 And BoardTo = 61) Then TriedToCastle = True
            If WhiteKingHasMoved = False And WhiteInCheck = False Then
                'Castle?
                If BoardFrom = 59 And BoardTo = 57 Then
                    LegalMove = True
                    If XMove = True Then
                        'Castle King Side
                        RJSoftChess.Board(58).Picture = RJSoftChess.Board(56).Picture
                        RJSoftChess.Board(58).Tag = RJSoftChess.Board(56).Tag
                        RJSoftChess.Board(56).Picture = RJSoftChess.Master_Blank.Picture
                        RJSoftChess.Board(56).Tag = ""
                        WhiteCanCastleKingSide = False
                        WhiteCanCastleQueenSide = False
                    End If
                End If
                If BoardFrom = 59 And BoardTo = 61 Then
                    LegalMove = True
                    If XMove = True Then
                        'Castle Queen Side
                        RJSoftChess.Board(60).Picture = RJSoftChess.Board(63).Picture
                        RJSoftChess.Board(60).Tag = RJSoftChess.Board(63).Tag
                        RJSoftChess.Board(63).Picture = RJSoftChess.Master_Blank.Picture
                        RJSoftChess.Board(63).Tag = ""
                        WhiteCanCastleKingSide = False
                        WhiteCanCastleQueenSide = False
                    End If
                End If
                If LegalMove = True And XMove = True Then WhiteKingHasMoved = True
            End If
        End If
     Else
        If ZRotated = True Then
            'Tried To Castle
            If (BoardFrom = 62 And BoardTo = 60) Or (BoardFrom = 60 And BoardTo = 58) Then TriedToCastle = True
            If BlackKingHasMoved = False And BlackInCheck = False Then
                'Castle?
                If BoardFrom = 60 And BoardTo = 62 Then
                    LegalMove = True
                    If XMove = True Then
                        'Castle King Side
                        RJSoftChess.Board(61).Picture = RJSoftChess.Board(63).Picture
                        RJSoftChess.Board(61).Tag = RJSoftChess.Board(63).Tag
                        RJSoftChess.Board(63).Picture = RJSoftChess.Master_Blank.Picture
                        RJSoftChess.Board(63).Tag = ""
                        BlackCanCastleKingSide = False
                        BlackCanCastleQueenSide = False
                    End If
                End If
                If BoardFrom = 60 And BoardTo = 58 Then
                    LegalMove = True
                    If XMove = True Then
                        'Castle Queen Side
                        RJSoftChess.Board(59).Picture = RJSoftChess.Board(56).Picture
                        RJSoftChess.Board(59).Tag = RJSoftChess.Board(56).Tag
                        RJSoftChess.Board(56).Picture = RJSoftChess.Master_Blank.Picture
                        RJSoftChess.Board(56).Tag = ""
                        BlackCanCastleKingSide = False
                        BlackCanCastleQueenSide = False
                    End If
                End If
                If LegalMove = True And XMove = True Then BlackKingHasMoved = True
            End If
         Else
            'Tried To Castle
            If (BoardFrom = 3 And BoardTo = 1) Or (BoardFrom = 3 And BoardTo = 5) Then TriedToCastle = True
            If BlackKingHasMoved = False And BlackInCheck = False Then
                'Castle?
                If BoardFrom = 3 And BoardTo = 1 Then
                    LegalMove = True
                    If XMove = True Then
                        'Castle King Side
                        RJSoftChess.Board(2).Picture = RJSoftChess.Board(0).Picture
                        RJSoftChess.Board(2).Tag = RJSoftChess.Board(0).Tag
                        RJSoftChess.Board(0).Picture = RJSoftChess.Master_Blank.Picture
                        RJSoftChess.Board(0).Tag = ""
                        BlackCanCastleKingSide = False
                        BlackCanCastleQueenSide = False
                    End If
                End If
                If BoardFrom = 3 And BoardTo = 5 Then
                    LegalMove = True
                    If XMove = True Then
                        'Castle Queen Side
                        RJSoftChess.Board(4).Picture = RJSoftChess.Board(7).Picture
                        RJSoftChess.Board(4).Tag = RJSoftChess.Board(7).Tag
                        RJSoftChess.Board(7).Picture = RJSoftChess.Master_Blank.Picture
                        RJSoftChess.Board(7).Tag = ""
                        BlackCanCastleKingSide = False
                        BlackCanCastleQueenSide = False
                    End If
                End If
                If LegalMove = True And XMove = True Then BlackKingHasMoved = True
            End If
        End If
    End If
    
    GoTo TestMove
    Exit Function
End If

Exit Function

TestMove:
    'Can not take your on piece
    If ZWhosTurn = CWhite Then
        If Left(RJSoftChess.Board(BoardTo).Tag, 1) = "W" Then LegalMove = False
     Else
        If Left(RJSoftChess.Board(BoardTo).Tag, 1) = "B" Then LegalMove = False
    End If
    If LegalMove = True And XMove = True Then RJSoftChess.Board(BoardTo).Picture = RJSoftChess.Master_Blank.Picture
Exit Function

TestAngle:
    'Can not jump over own piece
    'What Direction?
    ' 7 = Upper Left
    ' 9 = Upper Right
    '-7 = Lower Right
    '-9 = Lower Left
    ZDirection = ((BoardFrom - BoardTo) / Abs(Int(BoardTo / 8) - Int(BoardFrom / 8)))
    If ZDirection = 0 Then Return
    For XY = BoardFrom - ZDirection To BoardTo Step -ZDirection
        If ZWhosTurn = CWhite Then
            If Left(RJSoftChess.Board(XY).Tag, 1) = "W" Then
                LegalMove = False
                Exit For
            End If
            If (XY <> BoardTo) And (Left(RJSoftChess.Board(XY).Tag, 1) = "B") Then
                LegalMove = False
                Exit For
            End If
         Else
            If Left(RJSoftChess.Board(XY).Tag, 1) = "B" Then
                LegalMove = False
                Exit For
            End If
            If (XY <> BoardTo) And (Left(RJSoftChess.Board(XY).Tag, 1) = "W") Then
                LegalMove = False
                Exit For
            End If
        End If
    Next XY
Return

TestUpDownSideSide:
    If Abs(ZDirection) <> 8 Then
        If BoardFrom < BoardTo Then
            For XY = BoardFrom + 1 To BoardTo
                If ZWhosTurn = CWhite Then
                    If Left(RJSoftChess.Board(XY).Tag, 1) = "W" Then
                        LegalMove = False
                        Exit For
                    End If
                    If (XY <> BoardTo) And (Left(RJSoftChess.Board(XY).Tag, 1) = "B") Then
                        LegalMove = False
                        Exit For
                    End If
                 Else
                    If Left(RJSoftChess.Board(XY).Tag, 1) = "B" Then
                        LegalMove = False
                        Exit For
                    End If
                    If (XY <> BoardTo) And (Left(RJSoftChess.Board(XY).Tag, 1) = "W") Then
                        LegalMove = False
                        Exit For
                    End If
                End If
            Next XY
         Else
            For XY = BoardFrom - 1 To BoardTo Step -1
                If ZWhosTurn = CWhite Then
                    If Left(RJSoftChess.Board(XY).Tag, 1) = "W" Then
                        LegalMove = False
                        Exit For
                    End If
                    If (XY <> BoardTo) And (Left(RJSoftChess.Board(XY).Tag, 1) = "B") Then
                        LegalMove = False
                        Exit For
                    End If
                 Else
                    If Left(RJSoftChess.Board(XY).Tag, 1) = "B" Then
                        LegalMove = False
                        Exit For
                    End If
                    If (XY <> BoardTo) And (Left(RJSoftChess.Board(XY).Tag, 1) = "W") Then
                        LegalMove = False
                        Exit For
                    End If
                End If
            Next XY
        End If
    Else
        For XY = BoardFrom - ZDirection To BoardTo Step -ZDirection
            If ZWhosTurn = CWhite Then
                If Left(RJSoftChess.Board(XY).Tag, 1) = "W" Then
                    LegalMove = False
                    Exit For
                End If
                If (XY <> BoardTo) And (Left(RJSoftChess.Board(XY).Tag, 1) = "B") Then
                    LegalMove = False
                    Exit For
                End If
             Else
                If Left(RJSoftChess.Board(XY).Tag, 1) = "B" Then
                    LegalMove = False
                    Exit For
                End If
                If (XY <> BoardTo) And (Left(RJSoftChess.Board(XY).Tag, 1) = "W") Then
                    LegalMove = False
                    Exit For
                End If
            End If
        Next XY
    End If
    If XMove = True Then
        If ZWhosTurn = CWhite Then
            If ZRotated = True Then
                If WhiteKingHasMoved = False Then
                    If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = 7 Then WhiteCanCastleKingSide = False
                    If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = 0 Then WhiteCanCastleQueenSide = False
                    If WhiteCanCastleKingSide = False And WhiteCanCastleQueenSide = False Then WhiteKingHasMoved = True
                End If
             Else
                If WhiteKingHasMoved = False Then
                    If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = 56 Then WhiteCanCastleKingSide = False
                    If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = 63 Then WhiteCanCastleQueenSide = False
                    If WhiteCanCastleKingSide = False And WhiteCanCastleQueenSide = False Then WhiteKingHasMoved = True
                End If
            End If
         Else
            If ZRotated = True Then
                If BlackKingHasMoved = False Then
                    If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = 63 Then BlackCanCastleKingSide = False
                    If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = 56 Then BlackCanCastleQueenSide = False
                    If BlackCanCastleKingSide = False And BlackCanCastleQueenSide = False Then BlackKingHasMoved = True
                End If
             Else
                If BlackKingHasMoved = False Then
                    If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = 0 Then BlackCanCastleKingSide = False
                    If Abs(BoardFrom - (Int(BoardFrom / 8) * 8)) = 7 Then BlackCanCastleQueenSide = False
                    If BlackCanCastleKingSide = False And BlackCanCastleQueenSide = False Then BlackKingHasMoved = True
                End If
            End If
        End If
    End If
Return

End Function

Function CheckForCheck(CheckForAlreadyInCheck As Boolean)

Dim X As Integer, InCheck As Boolean, P As Integer
Dim WhereFrom As Integer, WhereTo As Integer, ZTempKing As String, ZTempLegalMove As Boolean
Dim KingLocation As Integer, ZTempWhatPiece As String, ZTempFromTag As String, ZTempToTag As String

InCheck = False
WhereFrom = BoardFrom
WhereTo = BoardTo
ZTempWhatPiece = WhatPiece
ZTempLegalMove = LegalMove

If CheckForAlreadyInCheck = True Then
    If ZWhosTurn = CWhite Then
        ZTempKing = "WKK"
        ZWhosTurn = CBlack
     Else
        ZTempKing = "BKK"
        ZWhosTurn = CWhite
    End If
 Else
    If ZWhosTurn = CWhite Then
        ZTempKing = "BKK"
     Else
        ZTempKing = "WKK"
    End If
End If

'King Location
If Right(ZTempWhatPiece, 2) = "KK" Then 'Check to see if it's the King that moved
    If CheckForAlreadyInCheck = True And TriedToCastle = True Then
        If BlackInCheck = True Or WhiteInCheck = True Then
            WhatPiece = ZTempWhatPiece
            BoardFrom = WhereFrom
            BoardTo = WhereTo
            LegalMove = False
            If ZWhosTurn = CWhite Then
                ZWhosTurn = CBlack
             Else
                ZWhosTurn = CWhite
            End If
            Exit Function
        End If
    End If
    KingLocation = BoardTo
 Else
    For X = 0 To 63
        If Trim(RJSoftChess.Board(X).Tag) = ZTempKing Then
            KingLocation = X
            Exit For
        End If
    Next X
End If

If CheckForAlreadyInCheck = True Then
    ZTempFromTag = RJSoftChess.Board(BoardFrom).Tag
    ZTempToTag = RJSoftChess.Board(BoardTo).Tag
    RJSoftChess.Board(BoardTo).Tag = RJSoftChess.Board(BoardFrom).Tag
    RJSoftChess.Board(BoardFrom).Tag = ""
End If

'Check all locations to see if king can be killed
For P = 0 To 63
    WhatPiece = Right(RJSoftChess.Board(P).Tag, 2)
    BoardFrom = P
    BoardTo = KingLocation
    If (Trim(Left(RJSoftChess.Board(P).Tag, 1)) <> Left(ZTempKing, 1)) Then
        CheckForLegalMove False
        If LegalMove = True Then
            If Trim(RJSoftChess.Board(BoardTo).Tag) = ZTempKing Then
                InCheck = True
                Exit For
            End If
        End If
    End If
Next P

WhatPiece = ZTempWhatPiece
BoardFrom = WhereFrom
BoardTo = WhereTo
LegalMove = ZTempLegalMove

If CheckForAlreadyInCheck = True Then
    RJSoftChess.Board(BoardFrom).Tag = ZTempFromTag
    RJSoftChess.Board(BoardTo).Tag = ZTempToTag
    If ZWhosTurn = CWhite Then
        ZWhosTurn = CBlack
     Else
        ZWhosTurn = CWhite
    End If
End If

BlackInCheck = False
WhiteInCheck = False
If CheckForAlreadyInCheck = True Then
    If InCheck = True Then
        If ZWhosTurn = CWhite Then
            BlackInCheck = Not InCheck
            WhiteInCheck = InCheck
         Else
            BlackInCheck = InCheck
            WhiteInCheck = Not InCheck
        End If
    End If
 Else
    If InCheck = True Then
        If ZWhosTurn = CWhite Then
            BlackInCheck = InCheck
            WhiteInCheck = Not InCheck
         Else
            BlackInCheck = Not InCheck
            WhiteInCheck = InCheck
        End If
    End If
End If

End Function

