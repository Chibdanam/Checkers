Attribute VB_Name = "Player"
'/// FONCTION   : Joue un tour comme si un humain interagissait avec le damier. Retourne true si un pion est d�plac�
'/// PARAM�TRE  : Range
'/// RETOUR     : Boolean
Public Function Run(pTarget As Range) As Boolean
Dim pawn As PawnModel
Dim enemyPawn As PawnModel
Dim checkerBoard As BoardModel
Dim pawnInitalState As PawnModel

    Set checkerBoard = New BoardModel
    Set pawn = New PawnModel
    Call pawn.Build(pTarget)
    
    'tant qu'aucun pion ne s'est d�plac�, c'est encore � nous de jouer
    Run = False
    
    'si la cellule cibl�e est un pion de notre couleur
    If pawn.IsPawn() And pawn.Color = checkerBoard.CurrentTurn Then
        
        'on m�morise le pion sur le plateau
        checkerBoard.Memory = pawn
    
    'si la cellule cibl�e est vide
    ElseIf Not pawn.IsPawn() Then
        
        'si un pion est m�moris� sur le plateau
        If checkerBoard.Memory.IsPawn() Then
            
            'on instancie le pion en m�moire
            Set pawnInitalState = checkerBoard.Memory
            
            'si le mouvement du pion m�moris� vers la cellule cibl�e est possible
            If pawnInitalState.TryMoveTo(pawn, True) Then
                'on efface le pion m�moris�
                Range("Memory").ClearContents
                'un pion s'est d�plac�
                Run = True
               
            End If
            
        End If
    
    End If
    
End Function