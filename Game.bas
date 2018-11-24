Attribute VB_Name = "Game"
'/// https://www.exceltrick.com/formulas_macros/vba-wait-and-sleep-functions/?
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'/// FONCTION   : Joue un tour comme si un humain interagissait avec la damier. Returne true si un pion est d�plac�
'/// PARAM�TRE  : Range
'/// RETOUR     : Boolean
Public Function Play(Target As Range) As Boolean

    Dim pawn As PawnModel
    Dim enemyPawn As PawnModel
    Dim checkerBoard As BoardModel

    Set checkerBoard = New BoardModel
    Set pawn = New PawnModel
    Call pawn.Build(Target)
    
    'tant qu'aucun pion ne s'est d�plac�, c'est encore a nous de jouer
    Play = False
    
    'si la cellule cibl� est un pion de notre couleur
    If pawn.IsPawn() And pawn.Color = checkerBoard.CurrentTurn Then
        
        'on m�morise le pion sur le plateau
        checkerBoard.Memory = pawn
    
    'si la cellule cibl� est vide
    ElseIf Not pawn.IsPawn() Then
        
        'si un pion est m�moris� sur le plateau
        If checkerBoard.Memory.IsPawn() Then
            
            'on instancie le pion en m�moire
            Dim pawnInitalState As PawnModel
            Set pawnInitalState = checkerBoard.Memory
            
            'si le mouvement du pion m�moris� vers la cellule cibl� est possible
            If pawnInitalState.TryMoveTo(pawn) Then
                'on efface le pion m�moris�
                Range("Memory").ClearContents
                'un pion s'est d�pac�
                Play = True
               
            End If
            
        End If
    
    End If
    
End Function


'/// FONCTION   : Retourne un tableau contenant tous les pions de la couleur pass�e en param�tre
'/// PARAM�TRE  : EColor
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
        
        'si le pion est v�ritablement un pion et est de la couleur pass� en param�tre
        If pawnCandidate.IsPawn And pawnCandidate.Color = pColor Then
            
            'on redimensionne notre tableau
            ReDim Preserve pawnList(pawnCounter)
            'on ajoute le pion au tableau de pion
            Set pawnList(pawnCounter) = pawnCandidate
            'on incr�mente notre compteur
            pawnCounter = pawnCounter + 1
            
        End If
        
    Next cell
    
    'on associe le tableau de pion ainsi constitu� au retour de la fonction
    GetPawns = pawnList
    
End Function


