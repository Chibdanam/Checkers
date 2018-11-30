Attribute VB_Name = "Player"
'/// FONCTION   : Joue un tour comme si un humain interagissait avec le damier. Retourne true si un pion est déplacé
'/// PARAMÈTRE  : Range
'/// RETOUR     : Boolean
Public Function Run(pTarget As Range) As Boolean
Dim pawn As PawnModel
Dim enemyPawn As PawnModel
Dim checkerBoard As BoardModel
Dim pawnInitalState As PawnModel

    Set checkerBoard = New BoardModel
    Set pawn = New PawnModel
    Call pawn.Build(pTarget)
    
    'tant qu'aucun pion ne s'est déplacé, c'est encore à nous de jouer
    Run = False
    
    'si la cellule ciblée est un pion de notre couleur
    If pawn.IsPawn() And pawn.Color = checkerBoard.CurrentTurn Then
        
        'on mémorise le pion sur le plateau
        checkerBoard.Memory = pawn
    
    'si la cellule ciblée est vide
    ElseIf Not pawn.IsPawn() Then
        
        'si un pion est mémorisé sur le plateau
        If checkerBoard.Memory.IsPawn() Then
            
            'on instancie le pion en mémoire
            Set pawnInitalState = checkerBoard.Memory
            
            'si le mouvement du pion mémorisé vers la cellule ciblée est possible
            If pawnInitalState.TryMoveTo(pawn, True) Then
                'on efface le pion mémorisé
                Range("Memory").ClearContents
                'un pion s'est déplacé
                Run = True
               
            End If
            
        End If
    
    End If
    
End Function