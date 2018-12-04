VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PawnModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private prvRow As Integer       '/// VARIABLE PRIVÉE : Ligne
Private prvColumn As Integer    '/// VARIABLE PRIVÉE : Colonne
Private prvColor As EColor      '/// VARIABLE PRIVÉE : Couleur
Private prvIsPawn As Boolean    '/// VARIABLE PRIVÉE : est véritablement un pion
Private prvIsQueen As Boolean   '/// VARIABLE PRIVÉE : est une dame



'/// CONSTRUCTEUR: méthode appelé lors de l'instanciation de l'objet
'/// PARAMÈTRE   : Aucun
'/// RETOUR      : Aucun
Private Sub class_Initialize()
    'on initialise les variables avec les valeurs de la cellule A1
    prvRow = 1
    prvColumn = 1
End Sub



'/// DESTRUCTEUR: méthode appelé lors de la destruction de l'objet
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Private Sub class_Terminate()
    
End Sub



'/// PROCÉDURE  : paramètre tous les attributs du pion à partir d'une range
'/// PARAMÈTRE  : Range
'/// RETOUR     : Aucun
Public Sub Build(ByVal cell As Range)
    
    prvRow = cell.Row
    prvColumn = cell.Column

    '/// On définit la variable privée de couleur du pion
    
    If cell.Font.Color = RGB(255, 255, 255) Then
        prvColor = EColor.White
    ElseIf cell.Font.Color = RGB(0, 0, 0) Then
        prvColor = EColor.Black
    End If
    
    '/// On définit si la cellule contient un pion
    
    If Not IsEmpty(cell) And (Me.Color = EColor.Black Or Me.Color = EColor.White) Then
        prvIsPawn = True
    Else
        prvIsPawn = False
    End If
    
    '/// On définit la variable privée d'état Dame du pion, si la valeur de la cellule est à alors c'est une dame
    
    If InStr(Chr(169), cell.Value) Then
        prvIsQueen = True
    Else
        prvIsQueen = False
    End If
    
End Sub



'/// PROPRIÉTÉ : Retourne la colonne du pion
'/// PARAMÈTRE : Aucun
'/// RETOUR    : Integer
Public Property Get Column() As Integer
    Column = prvColumn
End Property



'/// PROPRIÉTÉ : Modifie la variable privée correspondant à la colonne
'/// PARAMÈTRE : Integer
'/// RETOUR    : Aucun (Let)
Public Property Let Column(ByVal pColumn As Integer)
    prvColumn = pColumn
End Property



'/// PROPRIÉTÉ : Retourne la ligne du pion
'/// PARAMÈTRE : Aucun
'/// RETOUR    : Integer
Public Property Get Row() As Integer
    Row = prvRow
End Property



'/// PROPRIÉTÉ : Modifie la variable privée correspondant à la ligne
'/// PARAMÈTRE : Integer
'/// RETOUR    : Aucun (Let)
Public Property Let Row(ByVal pRow As Integer)
    prvRow = pRow
End Property



'/// PROPRIÉTÉ  : Retourne la valeur de la range du pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Range
Public Property Get CurrentRange() As Range
    Set CurrentRange = Cells(prvRow, prvColumn)
End Property



'/// PROPRIÉTÉ  : Modifie les variables privées correspondant à la ligne et à la colonne du pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Range
Public Property Let CurrentRange(ByVal pRange As Range)
    Me.Row = pRange.Row
    Me.Column = pRange.Column
End Property



'/// PROPRIÉTÉ  : Retourne la couleur du pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EColor
Public Property Get Color() As EColor
   Color = prvColor
End Property



'/// PROPRIÉTÉ  : Change la couleur du pion
'/// PARAMÈTRE  : EColor
'/// RETOUR     : Aucun
Public Property Let Color(ByVal pColor As EColor)
    
    prvColor = pColor
    
    Select Case pColor
        Case EColor.White
            CurrentRange.Font.Color = RGB(255, 255, 255)
        Case EColor.Black
            CurrentRange.Font.Color = RGB(0, 0, 0)
    End Select
    
End Property



'/// PROPRIÉTÉ  : Retourne la couleur des pions énnemis
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EColor
Public Property Get EnemyColor() As EColor
    
    If Me.Color = EColor.White Then
        EnemyColor = EColor.Black
    ElseIf Me.Color = EColor.Black Then
        EnemyColor = EColor.White
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne 1 si le pion est noir, -1 s'il est blanc
'///              Cela correspond à la direction selon Y (les lignes) vers laquelle le pion a le droit de se déplacer
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Integer
Public Property Get Orientation() As Integer
    
    If Me.Color = EColor.White Then
        Orientation = -1
    ElseIf Me.Color = EColor.Black Then
        Orientation = 1
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne vrai si le pion est une Dame
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get IsQueen() As Boolean
    IsQueen = prvIsQueen
End Property



'/// PROPRIÉTÉ  : Modifie la variable privée correspondant au rang du pion et le transforme en dame
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Let IsQueen(ByVal pIsQueen As Boolean)
    
    prvIsQueen = pIsQueen
    
    If pIsQueen Then
        CurrentRange.Value = Chr(169)
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne vrai si le pion est véritablement un pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get IsPawn() As Boolean
    IsPawn = prvIsPawn
End Property



'/// PROPRIÉTÉ  : Modifie la variable privée correspondant à l'etat du pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Let IsPawn(pIsPawn As Boolean)
    prvIsPawn = pIsPawn
End Property



'/// PROCÉDURE  : Transforme le pion en dame s'il est sur la dernière ligne de son orientation
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub TryBecomeQueen()

    If Not IsQueen Then
        If Me.Color = EColor.Black And Me.Row = 9 Or _
           Me.Color = EColor.White And Me.Row = 2 Then
            Me.IsQueen = True
        End If
    End If
    
End Sub



'/// FONCTION   : Vérifie si le mouvement est autorisé et, si possible, déplace le pion.
'///              Retourne vrai si le pion peut se déplacer
'/// PARAMÈTRE  : Pion (à la position finale après déplacement), Bouléen indiquant si on déplace le pion
'/// RETOUR     : Boolean
Public Function TryMoveTo(ByVal pTarget As PawnModel, ByVal pDoMove As Boolean) As Boolean
    
    If Me.CanMoveTo(pTarget) Then
        TryMoveTo = True
        If pDoMove Then
            Call MoveTo(pTarget)
        End If
    End If
    
End Function



'/// PROCÉDURE  : Déplace le pion vers la position passé en paramètre
'/// PARAMÈTRE  : Pion (é la position finale après déplacement)
'/// RETOUR     : Aucun
Public Sub MoveTo(ByVal pTarget As PawnModel)
Dim move As MoveModel

    Set move = New MoveModel
    Call move.Build(Me, pTarget)
    
    'si c'est une attaque
    If move.IsAttack Then
        'on efface le pion mangé
        move.EnemyKilled.CurrentRange.ClearContents
    End If
    
    'on efface le pion de sa position initial
    Me.CurrentRange.ClearContents
    'on associe la nouvelle range au pion
    Me.CurrentRange = pTarget.CurrentRange
    'on vérifie si le pion peut devenir une dame
    Call Me.TryBecomeQueen
    'on imprime le pion sur le plateau à sa nouvelle position
    Call Me.PrintPawnOnBoard
    
End Sub



'/// PROPRIÉTÉ  : Vérifie si le mouvement est autorisé
'/// PARAMÈTRE  : Pion (é la position finale après déplacement)
'/// RETOUR     : Boolean
Public Property Get CanMoveTo(ByVal pTarget As PawnModel) As Boolean
Dim move As MoveModel
    
    'on initialise le retour de la fonction a false
    CanMoveTo = False
    
    'on instancie un nouvel objet déplacement que l'on construit avec le pion actuel et le pion à la position finale souhaité
    Set move = New MoveModel
    Call move.Build(Me, pTarget)
    
    'si c'est un déplacement et aucun pion allié ne peut attaquer ou que ce déplacement est une attaque
    If (move.IsMove And Not Me.HasAttackingAlly) Or move.IsAttack Then
        CanMoveTo = True
    End If
    
    
End Property



'/// PROPRIÉTÉ  : Retourne true si le pion peut attaquer
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get CanAttack() As Boolean
Dim board As BoardModel
Dim targetCandidate As PawnModel
Dim move As MoveModel
Dim coefVector As Variant
Dim cardinalDirections As Variant
Dim cardinalDirection As Variant
Dim outOfRange As Boolean
Dim iVector As Integer
Dim nextRow As Integer
Dim nextColumn As Integer
Dim nextCell As Range
    
    'on initialise le retour de la fonction a false
    CanAttack = False
    
    Set targetCandidate = New PawnModel
    Set board = New BoardModel
    Set move = New MoveModel

    'si le pion est véritablement un pion
    If Me.IsPawn Then
    
        'on construit un tableau remplit des 4 directions possibles
        cardinalDirections = Array(EWindRose.NorthEast, EWindRose.NorthWest, EWindRose.SouthEast, EWindRose.SouthWest)
        
        'pour chaque direction possible dans le tableau des directions possible
        For Each cardinalDirection In cardinalDirections
        
            'on récupère les coefficients associés à la direction du déplacement
            coefVector = board.CardinalCoef(cardinalDirection)
            
            iVector = 0
            outOfRange = False
            
            'si le pion est une dame
            If IsQueen Then
            
                'Tant qu’on n’a pas détecter un pion à manger ou la fin du plateau, on boucle
                While Not (outOfRange)
                
                    iVector = iVector + 1
                    
                    nextRow = Me.Row + iVector * coefVector(0)
                    nextColumn = Me.Column + iVector * coefVector(1)
                    
                    'on récupère les coordonnées de la nième cellule sur la diagonale
                    Set nextCell = Cells(nextRow, nextColumn)
                    
                    'si cellule sur damier
                    If board.GetSection(nextCell) = ESection.Game Then
                    
                        'on construit l'objet pion
                        Set targetCandidate = New PawnModel
                        Call targetCandidate.Build(nextCell)
                        
                        'si la case est vide et est prête a accueillir un déplacement
                        If Not targetCandidate.IsPawn Then
                            
                            'on construit un objet déplacement à partir du pion actuel et la cellule en cours de test comme destination de mouvement
                            Call move.Build(Me, targetCandidate)
                            
                            'si le déplacement est une attaque
                            If move.IsAttack Then
                            
                                CanAttack = True
                                
                                Exit Property
                                
                            End If
                         End If
                    'si cellule hors du damier
                    Else
                        outOfRange = True
                    End If
                Wend
            'si le pion n'est pas une dame
            Else
                
                nextRow = Me.Row + 2 * coefVector(0)
                nextColumn = Me.Column + 2 * coefVector(1)
                
                If nextRow > 0 And nextColumn > 0 Then
                
                    Set nextCell = Cells(nextRow, nextColumn)
                    
                    'si cellule sur damier
                    If board.GetSection(nextCell) = ESection.Game Then
                
                        'on construit l'objet pion
                        Set targetCandidate = New PawnModel
                        Call targetCandidate.Build(nextCell)
                        
                        'si la case est vide et est prête à accueillir un déplacement
                        If Not targetCandidate.IsPawn Then
                            
                            'on construit un objet déplacement à partir du pion actuel et la cellule en cours de test comme destination de mouvement
                            Call move.Build(Me, targetCandidate)
                            
                            'si le déplacement est une attaque
                            If move.IsAttack Then
                            
                                CanAttack = True
                                
                                Exit Property
                                
                            End If
                         End If
                    End If
                End If
            End If
        Next cardinalDirection
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne true si le pion peut se déplacer
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get CanMove() As Boolean
Dim move As MoveModel
Dim coefVector As Variant
Dim cardinalDirections As Variant
Dim cardinalDirection As Variant
Dim Row As Integer, Column As Integer
Dim nextCell As Range
Dim targetCandidate As PawnModel
Dim board As BoardModel

    'on initialise le retour de la fonction a false
    CanMove = False
    
    'si aucun allié ne peut attaquer
    If Me.IsPawn And Not Me.HasAttackingAlly Then
    
        Set targetCandidate = New PawnModel
        Set board = New BoardModel
        Set move = New MoveModel
    
        'on construit un tableau remplit des 4 directions possibles
        cardinalDirections = Array(EWindRose.NorthEast, EWindRose.NorthWest, EWindRose.SouthEast, EWindRose.SouthWest)
        
        'pour chaque direction possible dans le tableau des directions possible
        For Each cardinalDirection In cardinalDirections
        
            'on récupère les coefficients associés à la direction du déplacement
            coefVector = board.CardinalCoef(cardinalDirection)
            
            'si le pion est une dame
            'ou si on est en train de tester une direction au nord pour un pion blanc, sud pour un pion noir
            If Me.IsQueen Or coefVector(0) = Me.Orientation Then
        
                Row = Me.Row + 1 * coefVector(0)
                Column = Me.Column + 1 * coefVector(1)
                
                'on récupère les coordonnées de la nième cellule sur la diagonale
                Set nextCell = Cells(Row, Column)
                
                'si cellule sur damier
                If board.GetSection(nextCell) = ESection.Game Then
                
                    'on construit l'objet pion
                    Set targetCandidate = New PawnModel
                    Call targetCandidate.Build(nextCell)
                    
                    'si la case est vide et est prête à accueillir un déplacement
                    If Not targetCandidate.IsPawn Then
                        
                        'on construit un objet déplacement à partir du pion actuel et la cellule en cours de test comme destination de mouvement
                        Call move.Build(Me, targetCandidate)
                        
                        'si le déplacement est un mouvement
                        If move.IsMove Then
                        
                            CanMove = True
                            
                            Exit Property
                            
                        End If
                    End If
                End If
            End If
        Next cardinalDirection
    End If
            
End Property



'/// PROPRIÉTÉ  : Retourne true si un allié peut attaquer
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get HasAttackingAlly() As Boolean
Dim alliesPawn As Variant
Dim allyPawn As Variant
    
    'Set currentPawn = New PawnModel
    
    HasAttackingAlly = False
    
    'on récupère le tableau des pions alliés
    alliesPawn = Tools.GetPawns(Me.Color)
    
    'pour chaque pion allié dans le tableau de pion alliés
    For Each allyPawn In alliesPawn
        
        'si le pion allié peut attaquer
        If allyPawn.CanAttack Then
        
            HasAttackingAlly = True
            Exit For
        End If
    Next allyPawn
End Property



'/// PROCÉDURE  : Écrit le pion sur le plateau
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub PrintPawnOnBoard()
        
    If Me.IsPawn And Me.IsQueen Then
        Me.CurrentRange.Value = Chr(169)
    ElseIf Me.IsPawn Then
        Me.CurrentRange.Value = "O"
    End If
    If Me.Color = EColor.White Then
        Me.CurrentRange.Font.Color = RGB(255, 255, 255)
    ElseIf Me.Color = EColor.Black Then
        Me.CurrentRange.Font.Color = RGB(0, 0, 0)
    End If
        
End Sub
