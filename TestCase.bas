Attribute VB_Name = "TestCase"
'/// PROCÉDURE  :
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun (sub)
Public Sub TestBecomeQueen()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "White"
    
    Range("B3").Value = "O"
    Range("D3").Value = "O"
    Range("F3").Value = "O"
    Range("H3").Value = "O"
    Range("B3").Font.ColorIndex = 2
    Range("D3").Font.ColorIndex = 2
    Range("F3").Font.ColorIndex = 2
    Range("H3").Font.ColorIndex = 2
    
    Range("C8").Value = "O"
    Range("E8").Value = "O"
    Range("G8").Value = "O"
    Range("I8").Value = "O"
    Range("C8").Font.ColorIndex = -4105
    Range("E8").Font.ColorIndex = -4105
    Range("G8").Font.ColorIndex = -4105
    Range("I8").Font.ColorIndex = -4105

End Sub



'/// PROCÉDURE  :
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun (sub)
Public Sub TestAttack()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "Black"

    Range("F5").Value = "O"
    Range("B8").Font.ColorIndex = -4105
    
    Range("G4").Value = "O"
    Range("G6").Value = "O"
    Range("E6").Value = "O"
    Range("G4").Font.ColorIndex = 2
    Range("G6").Font.ColorIndex = 2
    Range("E6").Font.ColorIndex = 2
    
End Sub



Public Sub TestQueenMove()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "White"
    
    Range("F3").Value = Chr(169)
    Range("E4").Value = Chr(169)
    Range("D5").Value = Chr(169)
    Range("C6").Value = Chr(169)
    Range("F3").Font.ColorIndex = 2
    Range("E4").Font.ColorIndex = 2
    Range("D5").Font.ColorIndex = 2
    Range("C6").Font.ColorIndex = 2
    
    Range("H5").Value = "O"
    Range("H7").Value = "O"
    Range("I8").Value = "O"
    Range("H5").Font.ColorIndex = -4105
    Range("H7").Font.ColorIndex = -4105
    Range("I8").Font.ColorIndex = -4105
    
    Range("G8").Value = "O"
    Range("G8").Font.ColorIndex = 2
    
End Sub

Public Sub TestYouWin()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "White"
    
    Range("F3").Value = "O"
    Range("E4").Value = "O"
    Range("F3").Font.ColorIndex = 2
    Range("E4").Font.ColorIndex = -4105
    
End Sub


Public Sub TestWeird()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "Black"
    
    Range("G4").Value = Chr(169)
    Range("H9").Value = "O"
    Range("G4").Font.ColorIndex = 2
    Range("H9").Font.ColorIndex = 2
    
    Range("D9").Value = Chr(169)
    Range("I8").Value = "O"
    Range("D9").Font.ColorIndex = -4105
    Range("I8").Font.ColorIndex = -4105
    
End Sub
