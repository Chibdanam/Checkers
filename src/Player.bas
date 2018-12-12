Attribute VB_Name = "Player"
Option Explicit
Option Base 0

'/// FONCTION   : Joue un tour comme si un humain interagissait avec le damier. Retourne true si un pion est d�plac�
'/// PARAM�TRE  : Range
'/// RETOUR     : Boolean
Public Function Run(ByVal pTarget As Range) As Boolean
Dim pawn As PawnModel
Dim enemyPawn As PawnModel
Dim board As BoardModel
Dim pawnInitalState As PawnModel

    Set board = New BoardModel
    Set pawn = New PawnModel
    Call pawn.Build(pTarget)
    
    'tant qu'aucun pion ne s'est d�plac�, c'est encore � nous de jouer
    Run = False
    
    'si la cellule cibl�e est un pion de notre couleur
    If pawn.IsPawn() And pawn.Color = board.TurnColor Then
        
        'on m�morise le pion sur le plateau
        board.Memory = pawn
    
    'si la cellule cibl�e est vide
    ElseIf Not pawn.IsPawn() Then
        
        'si un pion est m�moris� sur le plateau
        If board.Memory.IsPawn() Then
            
            'on instancie le pion en m�moire
            Set pawnInitalState = board.Memory
            
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
