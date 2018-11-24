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
    Dim pawnStepList As Variant

    Set checkerBoard = New BoardModel
    Set pawn = New PawnModel
    Call pawn.Build(Target)
    
    'tant qu'aucun pion ne s'est d�plac�, c'est encore a nous de jouer
    Play = False
    
    'si la cellule cibl� est un pion de notre couleur
    If pawn.IsPawn() And pawn.Color = checkerBoard.CurrentTurn Then
        
        Call checkerBoard.CleanMemory
        
        'on m�morise le pion sur le plateau
        checkerBoard.Memory = pawn
    
    'si la cellule cibl� est vide
    ElseIf Not pawn.IsPawn() Then
        
        pawnStepList = checkerBoard.PawnAndMovesFromMemory
        
        Dim pawnInitalState As PawnModel
        
        'si au moins un pion est m�moris� sur le plateau
        If Not IsArrayNullOrEmpty(pawnStepList) Then
            
            'si un seul pion est m�moris�
            If UBound(pawnStepList) = 0 Then
                
                'on instancie le pion en m�moire
                Set pawnInitalState = pawnStepList(0)
    
                'si le mouvement du pion m�moris� vers la cellule cibl� est possible
                If pawnInitalState.IsPawn() And _
                   pawnInitalState.Color = checkerBoard.CurrentTurn And _
                   pawnInitalState.TryMoveTo(pawn, True) Then
                    'on efface le pion m�moris�
                   Call checkerBoard.CleanMemory
                    'un pion s'est d�pac�
                    Play = True
    
                End If
            
            'si plusieurs pions sont m�moris�s
            ElseIf UBound(pawnStepList) > 0 Then
            
                Dim moveStepCount As Integer
                moveStepCount = 0
                
                Dim pawnFinalState As PawnModel
                Dim finalMove As Boolean
                Dim doMoves As Boolean
                
                Play = True
                finalMove = False
                doMoves = True
                
                'tant qu'on a pas effectuer la totalit� des d�placements m�moris�s
                While Not finalMove
                    
                    'on initialise le pion a d�placer et la cible de son d�placement
                    Set pawnInitalState = pawnStepList(moveStepCount)
                    Set pawnFinalState = pawnStepList(moveStepCount + 1)
                    
                    'si la cible du mouvement en cours est la cellule o� l'on a double cliqu�
                    If pawnFinalState.CurrentRange.Address = pawn.CurrentRange.Address Then
                        'alors ce d�placement est le dernier a effectuer
                        finalMove = True
                    End If
                    
                    If (pawnInitalState.IsPawn() Or pawnInitalState.IsStepMove()) And pawnFinalState.IsStepMove() Then
                    
                        If Not pawnInitalState.TryMoveTo(pawnFinalState, False) Then
                        
                            doMoves = False
                            
                        End If
                        
                    End If
                    
                    moveStepCount = moveStepCount + 1
                    
                Wend
                
                If doMoves Then
                
                    finalMove = False
                    moveStepCount = 0
                    
                    While Not finalMove
                    
                        'on initialise le pion a d�placer et la cible de son d�placement
                        Set pawnInitalState = pawnStepList(moveStepCount)
                        Set pawnFinalState = pawnStepList(moveStepCount + 1)
                        
                        'si la cible du mouvement en cours est la cellule o� l'on a double cliqu�
                        If pawnFinalState.CurrentRange.Address = pawn.CurrentRange.Address Then
                            'alors ce d�placement est le dernier a effectuer
                            finalMove = True
                        End If
                        
                        If (pawnInitalState.IsPawn() Or pawnInitalState.IsStepMove()) And pawnFinalState.IsStepMove() Then
                        
                            If pawnInitalState.TryMoveTo(pawnFinalState, True) Then
                                
                                'un pion s'est d�pac�
                                Play = True
                                
                            End If
                            
                        End If
                        
                        moveStepCount = moveStepCount + 1
                    
                    Wend
                
                        
                    'on efface le pion m�moris�
                    Call checkerBoard.CleanMemory(Play)
                            
                    
                End If
                
                
                
        
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


Function IsArrayNullOrEmpty(arr As Variant) As Boolean
    
    On Error Resume Next
    
    IsArrayNullOrEmpty = True
    
    If IsArray(arr) And _
    Not IsError(LBound(arr, 1)) And _
    LBound(arr, 1) <= UBound(arr, 1) And _
    Not IsEmpty(arr) And _
    UBound(arr) > -1 Then
        IsArrayNullOrEmpty = False
    End If
    
    Exit Function
    
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

