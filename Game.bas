Attribute VB_Name = "Game"
'/// https://www.exceltrick.com/formulas_macros/vba-wait-and-sleep-functions/?
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'/// FONCTION   : Joue un tour comme si un humain interagissait avec la damier. Returne true si un pion est déplacé
'/// PARAMÈTRE  : Range
'/// RETOUR     : Boolean
Public Function Play(Target As Range) As Boolean

    Dim pawn As PawnModel
    Dim enemyPawn As PawnModel
    Dim checkerBoard As BoardModel

    Set checkerBoard = New BoardModel
    Set pawn = New PawnModel
    Call pawn.Build(Target)
    
    'tant qu'aucun pion ne s'est déplacé, c'est encore a nous de jouer
    Play = False
    
    'si la cellule ciblé est un pion de notre couleur
    If pawn.IsPawn() And pawn.Color = checkerBoard.CurrentTurn Then
        
        'on mémorise le pion sur le plateau
        checkerBoard.Memory = pawn
    
    'si la cellule ciblé est vide
    ElseIf Not pawn.IsPawn() Then
        
        'si un pion est mémorisé sur le plateau
        If checkerBoard.Memory.IsPawn() Then
            
            'on instancie le pion en mémoire
            Dim pawnInitalState As PawnModel
            Set pawnInitalState = checkerBoard.Memory
            
            'si le mouvement du pion mémorisé vers la cellule ciblé est possible
            If pawnInitalState.TryMoveTo(pawn) Then
                'on efface le pion mémorisé
                Range("Memory").ClearContents
                'un pion s'est dépacé
                Play = True
               
            End If
            
        End If
    
    End If
    
End Function


'/// FONCTION   : Retourne un tableau contenant tous les pions de la couleur passée en paramètre
'/// PARAMÈTRE  : EColor
'/// RETOUR     : Variant (tableau de pion)
Public Function GetPawns(pColor As EColor) As Variant

    Dim pawnList() As PawnModel
    Dim pawnCounter As Integer
    pawnCounter = 0
    
    'pour chaque cellule du damier
    For Each cell In Range("Game")
        
        'on instancie un objet pion et on le construit avec la range de la cellule actuelle
        Set pawnCandidate = New PawnModel
        Call pawnCandidate.Build(cell.Cells(1, 1))
        
        'si le pion est véritablement un pion et est de la couleur passé en paramètre
        If pawnCandidate.IsPawn And pawnCandidate.Color = pColor Then
            
            'on redimensionne notre tableau
            ReDim Preserve pawnList(pawnCounter)
            'on ajoute le pion au tableau de pion
            Set pawnList(pawnCounter) = pawnCandidate
            'on incrémente notre compteur
            pawnCounter = pawnCounter + 1
            
        End If
        
    Next cell
    
    'on associe le tableau de pion ainsi constitué au retour de la fonction
    GetPawns = pawnList
    
End Function


