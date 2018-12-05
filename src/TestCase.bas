Attribute VB_Name = "TestCase"
Option Explicit
Option Base 0

Public Sub TestFormatBoard()
Dim board As BoardModel
    Set board = New BoardModel
    Call board.FormatBoard
    Call board.Initialisation
End Sub

Public Sub TestVoid()
Dim blueprint As String

    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestStart()
Dim blueprint As String

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|" + vbNewLine + _
                "|-|w|-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestBecomeQueen()
Dim blueprint As String
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-|w|-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-|b|-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestAttack()
Dim blueprint As String
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-|b|-|b|-| |" + vbNewLine + _
                "| |-| |-|w|-| |-|" + vbNewLine + _
                "|-| |-|b|-| |-| |" + vbNewLine + _
                "| |-| |-| |-|w|-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestQueenMove()
Dim blueprint As String
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| B|-|B|-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-|B|-|b|-|b|-|" + vbNewLine + _
                "|-|B|-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-|b|-|" + vbNewLine + _
                "|-| |-| |-|w|-| |" + vbNewLine + _
                "| |-| |-|b|-| |-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestYouWin()
Dim blueprint As String
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-|w|-| |-|" + vbNewLine + _
                "|-| |-|b|-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestWeird()
Dim blueprint As String

    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-|B|-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-|b|" + vbNewLine + _
                "| |-|B|-| |-|w|-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestSituation()
Dim blueprint As String

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-| |-| |-|b|-|b|" + vbNewLine + _
                "|b|-|b|-| |-| |-|" + vbNewLine + _
                "|-|w|-| |-| |-|w|" + vbNewLine + _
                "|w|-| |-|w|-| |-|" + vbNewLine + _
                "|-|w|-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestSituationOK()
Dim blueprint As String

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-| |-|w|-|b|-|b|" + vbNewLine + _
                "|b|-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-|w|" + vbNewLine + _
                "|w|-| |-|w|-| |-|" + vbNewLine + _
                "|-|w|-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestSituationKO()
Dim blueprint As String

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-| |-| |-|b|-|b|" + vbNewLine + _
                "|b|-|b|-| |-| |-|" + vbNewLine + _
                "|-|w|-| |-| |-|w|" + vbNewLine + _
                "|w|-|w|-|w|-| |-|" + vbNewLine + _
                "|-| |-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Tools.Compute(blueprint)
End Sub

Public Sub TestSituationBlocked()
Dim blueprint As String
    Range("TurnValue") = "Black"
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-|b|-|b|" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-|w|" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Tools.Compute(blueprint)
End Sub
