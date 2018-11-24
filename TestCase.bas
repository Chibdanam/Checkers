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
    Range("B3").Font.Color = RGB(255, 255, 255)
    Range("D3").Font.Color = RGB(255, 255, 255)
    Range("F3").Font.Color = RGB(255, 255, 255)
    Range("H3").Font.Color = RGB(255, 255, 255)
    
    Range("C8").Value = "O"
    Range("E8").Value = "O"
    Range("G8").Value = "O"
    Range("I8").Value = "O"
    Range("C8").Font.Color = RGB(0, 0, 0)
    Range("E8").Font.Color = RGB(0, 0, 0)
    Range("G8").Font.Color = RGB(0, 0, 0)
    Range("I8").Font.Color = RGB(0, 0, 0)

End Sub



'/// PROCÉDURE  :
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun (sub)
Public Sub TestAttack()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "Black"

    Range("F5").Value = "O"
    Range("B8").Font.Color = RGB(0, 0, 0)
    
    Range("G4").Value = "O"
    Range("G6").Value = "O"
    Range("E6").Value = "O"
    Range("G4").Font.Color = RGB(255, 255, 255)
    Range("G6").Font.Color = RGB(255, 255, 255)
    Range("E6").Font.Color = RGB(255, 255, 255)
    
End Sub



Public Sub TestQueenMove()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "White"
    
    Range("F3").Value = Chr(169)
    Range("E4").Value = Chr(169)
    Range("D5").Value = Chr(169)
    Range("C6").Value = Chr(169)
    Range("F3").Font.Color = RGB(255, 255, 255)
    Range("E4").Font.Color = RGB(255, 255, 255)
    Range("D5").Font.Color = RGB(255, 255, 255)
    Range("C6").Font.Color = RGB(255, 255, 255)
    
    Range("H5").Value = "O"
    Range("H7").Value = "O"
    Range("I8").Value = "O"
    Range("H5").Font.Color = RGB(0, 0, 0)
    Range("H7").Font.Color = RGB(0, 0, 0)
    Range("I8").Font.Color = RGB(0, 0, 0)
    
    Range("G8").Value = "O"
    Range("G8").Font.Color = RGB(255, 255, 255)
    
End Sub

Public Sub TestYouWin()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "White"
    
    Range("F3").Value = "O"
    Range("E4").Value = "O"
    Range("F3").Font.Color = RGB(255, 255, 255)
    Range("E4").Font.Color = RGB(0, 0, 0)
    
End Sub


Public Sub TestWeird()

    Range("Game").ClearContents
    Range("Memory").ClearContents
    Range("TurnValue") = "Black"
    
    Range("G4").Value = Chr(169)
    Range("H9").Value = "O"
    Range("G4").Font.Color = RGB(255, 255, 255)
    Range("H9").Font.Color = RGB(255, 255, 255)
    
    Range("D9").Value = Chr(169)
    Range("I8").Value = "O"
    Range("D9").Font.Color = RGB(0, 0, 0)
    Range("I8").Font.Color = RGB(0, 0, 0)
    
End Sub
