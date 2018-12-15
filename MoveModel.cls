VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MoveModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/// VARIABLE PRIVÉE : Pion que l'on veut déplacer
Private prvPawn As PawnModel
'/// VARIABLE PRIVÉE : Cible du déplacement
Private prvTarget As PawnModel
'/// VARIABLE PRIVÉE : Enemi tué lors du déplacement
Private prvEnemyKilled As PawnModel

'/// j'ai modif ton code =)

'/// CONSTRUCTEUR: methode appelé lors de l'instanciation de l'objet
'/// PARAMÈTRE   : Aucun (evenement constructeur)
'/// RETOUR      : Aucun (evenement)
Private Sub class_Initialize()
    
    'on instancie les variables privées pour pouvoir les manipuler dans les propriétés/fonctions/procédures
    Set prvPawn = New PawnModel
    Set prvTarget = New PawnModel
    Set prvEnemyKilled = New PawnModel
    
End Sub



'/// DESTRUCTEUR: methode appelé lors de la destruction de l'objet
'/// PARAMÈTRE  : Aucun (evenement destructeur)
'/// RETOUR     : Aucun (evenemnt)
Private Sub class_Terminate()

    'on vide la mémoire associée aux variables privées
    Set prvPawn = Nothing
    Set prvTarget = Nothing
    Set prvEnemyKilled = Nothing
    
End Sub



'/// PROCÉDURE  : Permet d'associer des valeurs aux variables privées
'/// PARAMÈTRES : Pion , Pion état final
'/// RETOUR     : Aucun
Public Sub Build(pawn As PawnModel, Target As PawnModel)
    
    'on associe des valeurs aux variables privées de notre objet 'move'
    Set prvPawn = pawn
    Set prvTarget = Target
    
End Sub



'/// PROPRIÉTÉ  : Récupère le déplacement suivant Y du mouvement (la différence de ligne entre le pion et la cible)
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Integer
Property Get Yvector() As Integer
    Yvector = prvTarget.Row - prvPawn.Row
End Property



'/// PROPRIÉTÉ  : Récupère le déplacement suivant X du mouvement (la différence de colonne entre le pion et la cible)
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Integer
Property Get Xvector() As Integer
    Xvector = prvTarget.Column - prvPawn.Column
End Property



'/// PROPRIÉTÉ  : Récupère la norme de déplacement du pion. La valeur sera 0 lorsque
'///              la cible ne sera pas positionnée diagonalement par rapport au pion.
'///              Cette valeur représente le nombre de case sur la diagonale séparant
'///              la position initiale et finale du pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Integer
Property Get Norme() As Integer
    
    'on initialise la norme a 0
    Norme = 0
    
    'si la valeur absolue du déplacement sur x est égale a la valeur absolue du déplacement sur y
    If Abs(Yvector) = Abs(Xvector) Then
        'alors le déplacement est sur une diagonale du pion
        Norme = Abs(Xvector)
    End If
    
End Property


'/// PROPRIÉTÉ  : Récupère l'orientation du mouvement par rapport a son emplacement initiale
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : La diagonale du mouvement
Property Get Direction() As EWindRose
    
    'si le signe du déplacement suivant X et suivant Y est positif
    If Sgn(Xvector) = 1 And Sgn(Yvector) = 1 Then
        'Alors le pion se déplace en direction du Sud Est (en bas à droite)
        Direction = EWindRose.SouthEast
        
    'si le signe du déplacement est positif suivant X et négatif suivant Y
    ElseIf Sgn(Xvector) = 1 And Sgn(Yvector) = -1 Then
        'Alors le pion se déplace en direction du Nord Est (en haut à droite)
        Direction = EWindRose.NorthEast
        
    'si le signe du déplacement est négatif suivant X et positif suivant Y
    ElseIf Sgn(Xvector) = -1 And Sgn(Yvector) = 1 Then
        'Alors le pion se déplace en direction du Sud Ouest (en bas à gauche)
        Direction = EWindRose.SouthWest
        
    'si le signe du déplacement suivant X et suivant Y est negatif
    ElseIf Sgn(Xvector) = -1 And Sgn(Yvector) = -1 Then
        'Alors le pion se déplace en direction du Nord Ouest (en haut à gauche)
        Direction = EWindRose.NorthWest
    End If
    
End Property



'/// PROPRIÉTÉ  : Récupère le pion tué lors du déplacement (Nothing si aucun pion n'est mangé pendant le déplacement)
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Pion
Property Get EnemyKilled() As PawnModel
    
    Set EnemyKilled = prvEnemyKilled
    
End Property



'/// PROPRIÉTÉ  : Retourne true si le déplacement correspond a un mouvement ou a une attaque
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Booléen
Property Get IsMoveOrAttack() As Boolean
    
    'on initialise le retour de la fonction a false
    IsMoveOrAttack = False
    
    'si le déplacement est un mouvement ou une attaque
    If Me.IsMove() Or Me.IsAttack() Then
    
        IsMoveOrAttack = True
        
    End If
        
       
End Property



'/// PROPRIÉTÉ  : Retourne true si le déplacement correspond a un mouvement
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Booléen
Property Get IsMove() As Boolean
    
    'on initialise le retour a false
    IsMove = False

    'si le pion est une dame
    If prvPawn.IsQueen Then
        
        'si le mouvement respect les régles du mouvement d'une dame
        If IsQueenMove() Then
            IsMove = True
        End If
        
    'si le pion n'est pas une dame
    Else
        
        'si le mouvement respect les régles du mouvement d'un pion
        If IsPawnMove() Then
            IsMove = True
        End If
        
    End If
        
       
End Property



'/// PROPRIÉTÉ  : Retourne true si le déplacement correspond a une attaque
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Booléen
Property Get IsAttack() As Boolean
    
    'on initialise le retour a false
    IsAttack = False
    
    'si le pion est une dame
    If prvPawn.IsQueen Then
    
        'si le mouvement respect les régles d'attaque d'une dame
        If IsQueenAttack() Then
            IsAttack = True
        End If
        
    'si le pion n'est pas une dame
    Else
        
        'si le mouvement respect les régles d'attaque d'un pion
        If IsPawnAttack() Then
            IsAttack = True
        End If
        
    End If
       
End Property



'/// PROPRIÉTÉ  : Retourne true si le déplacement respecte les régles d'un mouvement de pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Booléen
Property Get IsPawnMove() As Boolean
    
    'on initialise le retour a false
    IsPawnMove = False
    
    'si le déplacement n'est que d'une case,
    'et si le déplacement suivant Y est positif ( vers le bas ) si le pion est Noir, ou Négatif (vers le haut) si le pion est blanc
    If Norme = 1 And Yvector = prvPawn.Orientation Then
        IsPawnMove = True
    End If
       
End Property



'/// PROPRIÉTÉ  : Retourne true si le déplacement respecte les régles d'une attaque de pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Booléen
Property Get IsPawnAttack() As Boolean
    
    'on initialise le retour a false
    IsPawnAttack = False
    
    'si le déplacement est de deux cases
    If Norme = 2 Then
        
        'on instancie un objet pion pour représenter le pion attaqué
        Dim enemyPawn As PawnModel
        Set enemyPawn = New PawnModel
        
        'on récupere la cellule entre la position initiale et la position finale du pion
        Dim enemyCell As Range
        Set enemyCell = Cells(prvTarget.Row - Sgn(Yvector) * 1, prvTarget.Column - Sgn(Xvector) * 1)
        
        'on construit un objet pion a partir de la cellule
        Call enemyPawn.Build(enemyCell)
        
        'si l'objet pion est véritablement un pion, et s'il n'est pas de la couleur du pion attaquant
        If enemyPawn.IsPawn And enemyPawn.Color = prvPawn.EnemyColor Then
        
            'on garde les informations du pion tué dans la variable privée
            Set prvEnemyKilled = enemyPawn
            IsPawnAttack = True
            
        End If
        
    End If
       
End Property



'/// PROPRIÉTÉ  : Retourne true si le déplacement respecte les régles de mouvement d'une Dame
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Booléen
Property Get IsQueenMove() As Boolean
    
    Dim outOfRange As Boolean
    Dim maxVector As Integer
    
    'représente un tableau de (x,y) représentant les coefficient a appliqué pour ciblé les cellules d'une diagonale
    Dim coefVector As Variant
    
    Dim nextCell As Range
    Dim pawnCandidate As PawnModel
    Dim board As BoardModel
    
    'on initialise le retour a false
    IsQueenMove = False
    
    'si la déplacement est sur la diagonale
    If Norme <> 0 Then
    
        'on instancie un objet pion
        Set pawnCandidate = New PawnModel
        'on instancie un objet plateau de jeu
        Set board = New BoardModel
        
        'on récupere les coeeficients associés à la direction du déplacement
        coefVector = board.CardinalCoef(Me.Direction)
        'on initialise le déplacement maximum possible a 0
        maxVector = 0
        
        outOfRange = False
        
        'Tant qu'on a pas détécter un pion a manger ou la fin du plateau, on boucle
        While Not outOfRange
            
            'on incrémente de 1 la cellule maximal finale pour un potentiel déplacement
            maxVector = maxVector + 1
            
            'on récupère les coordonnées de la n-ieme cellule sur la diagonale
            Set nextCell = Cells(prvPawn.Row + maxVector * coefVector(0), prvPawn.Column + maxVector * coefVector(1))
            
            'si cellule sur damier
            If board.GetSection(nextCell) = ESection.Game Then
            
                'on construit l'objet pion
                Call pawnCandidate.Build(nextCell)
                
                'si le pion existe
                If pawnCandidate.IsPawn Then
    
                    'cette case est la limite de notre déplacement
                    outOfRange = True
                        
                End If
                                
            'si cellule en dehors du damier
            Else
                outOfRange = True
                
            End If
            
        Wend
    
        'si le déplacement respecte la limite trouvée
        If (Me.Norme < maxVector) Then
        
            'le mouvement est accepté
            IsQueenMove = True
        End If
    
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne true si le déplacement respecte les régles d'une attaque d'une dame
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Booléen
Property Get IsQueenAttack() As Boolean
        
    Dim outOfRange As Boolean
    Dim iVector As Integer
    Dim minVector As Integer
    Dim maxVector As Integer
    
    'représente un tableau de (x,y) représentant les coefficient a appliqué pour ciblé les cellules d'une diagonale
    Dim coefVector As Variant
    
    Dim nextCell As Range
    Dim pawnCandidate As PawnModel
    Dim pawnEnemy As PawnModel
    Dim board As BoardModel
    
    Set pawnCandidate = New PawnModel
    Set pawnEnemy = New PawnModel
    Set board = New BoardModel
    
    'on initialise le retour a false
    IsQueenAttack = False
    
    'si le déplacement est sur la diagonale
    If Norme <> 0 Then
    
        'on récupere les coeeficients associés à la direction du déplacement
        coefVector = board.CardinalCoef(Me.Direction)
        'on initialise le déplacement minimum possible a 0
        minVector = 0
        'on initialise le déplacement maximum possible a 0
        maxVector = 0
        
        iVector = 0
        
        outOfRange = False
        
        'Tant qu'on a pas détécter un pion a manger ou la fin du plateau, on boucle
        While Not outOfRange
            
            'on incrémente le vecteur représentant la n-eme cellule sur la diagonal
            iVector = iVector + 1
            
            'on récupère les coordonnées de la n-ieme cellule sur la diagonale
            Set nextCell = Cells(prvPawn.Row + iVector * coefVector(0), prvPawn.Column + iVector * coefVector(1))
            
            'si cellule sur damier
            If board.GetSection(nextCell) = ESection.Game Then
            
                'on construit l'objet pion
                Call pawnCandidate.Build(nextCell)
                
                'si le pion existe
                If pawnCandidate.IsPawn Then
                    
                    'si un pion enemi est mémorisé
                    If pawnEnemy.IsPawn Then
                        
                        'alors deux pions sont détecté et on ne peut se déplacer qu'entre eux, cette case est la limite de notre déplacement
                        outOfRange = True
                        maxVector = iVector
                        
                    'si aucun pion n'est mémorisé
                    Else
                    
                        'si le pion détécté est un pion enemis
                        If pawnCandidate.Color = prvPawn.EnemyColor Then
                            
                            'on l'enregistre
                            pawnEnemy.Row = pawnCandidate.Row
                            pawnEnemy.Column = pawnCandidate.Column
                            Call pawnEnemy.Build(pawnEnemy.CurrentRange)
                            
                            'notre déplacement commence a partir de cette case
                            minVector = iVector
                        
                        'si c'est un pion de notre couleur
                        Else
                            
                            'alors aucune attaque n'est possible
                            minVector = iVector
                            maxVector = iVector
                            outOfRange = True
                            
                        End If
                            
                    End If
                        
                'si il n'y pas de pion sur cette case
                Else
                
                    'si un pion enemi est mémorisé
                    If pawnEnemy.IsPawn Then
                    
                        'on considere que la cellule actuelle est un emplacement valide de déplacement apres attaque
                        maxVector = iVector
                    
                    End If
                    
                End If
                                
            'si cellule en dehors du damier
            Else
                outOfRange = True
                maxVector = iVector
            End If
            
        Wend
        
        'si le déplacement est compris entre les limites trouvé
        If (Me.Norme < maxVector) And (Me.Norme > minVector) And pawnEnemy.IsPawn Then
            
            'l'attaque est valide
            IsQueenAttack = True
            
            'on enregistre le pion capturé
            prvEnemyKilled.Row = pawnEnemy.Row
            prvEnemyKilled.Column = pawnEnemy.Column
            
        End If
        
    End If
            
End Property


