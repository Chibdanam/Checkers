Attribute VB_Name = "TestCase"
Public Sub TestVoid()

    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestStart()

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|" + vbNewLine + _
                "|-|w|-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestBecomeQueen()
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-|w|-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-|b|-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestAttack()
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-|b|-|b|-| |" + vbNewLine + _
                "| |-| |-|w|-| |-|" + vbNewLine + _
                "|-| |-|b|-| |-| |" + vbNewLine + _
                "| |-| |-| |-|w|-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestQueenMove()
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-|B|-| |-|" + vbNewLine + _
                "|-| |-|B|-| |-| |" + vbNewLine + _
                "| |-|B|-| |-|b|-|" + vbNewLine + _
                "|-|B|-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-|b|-|" + vbNewLine + _
                "|-| |-| |-|w|-|b|" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestYouWin()
    
    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-|w|-| |-|" + vbNewLine + _
                "|-| |-|b|-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestWeird()

    '             a b c d e f g h
    blueprint = "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-|B|-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-|b|" + vbNewLine + _
                "| |-|B|-| |-|w|-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestSituation()

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-| |-| |-|b|-|b|" + vbNewLine + _
                "|b|-|b|-| |-| |-|" + vbNewLine + _
                "|-|w|-| |-| |-|w|" + vbNewLine + _
                "|w|-| |-|w|-| |-|" + vbNewLine + _
                "|-|w|-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestSituationOK()

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-| |-|w|-|b|-|b|" + vbNewLine + _
                "|b|-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-|w|" + vbNewLine + _
                "|w|-| |-|w|-| |-|" + vbNewLine + _
                "|-|w|-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Compute(blueprint)
End Sub

Public Sub TestSituationKO()

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-| |-| |-|b|-|b|" + vbNewLine + _
                "|b|-|b|-| |-| |-|" + vbNewLine + _
                "|-|w|-| |-| |-|w|" + vbNewLine + _
                "|w|-|w|-|w|-| |-|" + vbNewLine + _
                "|-| |-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Compute(blueprint)
End Sub