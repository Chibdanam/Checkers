VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PawnModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/// VARIABLE PRIVÉE : Ligne
Private prvRow As Integer
'/// VARIABLE PRIVÉE : Colonne
Private prvColumn As Integer
'/// VARIABLE PRIVÉE : Couleur
Private prvColor As EColor
'/// VARIABLE PRIVÉE : Couleur
Private prvIsPawn As Boolean
'/// VARIABLE PRIVÉE : Colonne
Private prvIsQueen As Boolean



'/// CONSTRUCTEUR: methode appelé lors de l'instanciation de l'objet
'/// PARAMÈTRE   : Aucun (evenement constructeur)
'/// RETOUR      : Aucun (evenement)
Private Sub class_Initialize()
    'on initialise les variables avec les valeurs de la cellule A1
    prvRow = 1
    prvColumn = 1
End Sub



'/// DESTRUCTEUR: methode appelé lors de la destruction de l'objet
'/// PARAMÈTRE  : Aucun (evenement destructeur)
'/// RETOUR     : Aucun (evenemnt)
Private Sub class_Terminate()
    
End Sub



'/// PROCÉDURE  : parametre tous les attributs du pion à partir d'une range
'/// PARAMÈTRE  : Range
'/// RETOUR     : Aucun
Public Sub Build(cell As Range)
    
    prvRow = cell.Row
    prvColumn = cell.Column
    
    '/// On définit la variable privé de couleur du pion
    
    If cell.Font.Color = RGB(255, 255, 255) Then
        prvColor = EColor.White
    ElseIf cell.Font.Color = RGB(0, 0, 0) Then
        prvColor = EColor.Black
    ElseIf cell.Font.Color = RGB(150, 50, 50) Then
        prvColor = EColor.BlackMoveStep
    ElseIf cell.Font.Color = RGB(230, 180, 180) Then
        prvColor = EColor.WhiteMoveStep
    End If
    
    '/// On définit si la cellule contient un pion
    
    If Not IsEmpty(cell) And (Me.Color = EColor.Black Or Me.Color = EColor.White) Then
        prvIsPawn = True
    Else
        prvIsPawn = False
    End If
    
    '/// On définit la variable privé d'état Dame du pion, si la valeur de la cellule est © alors c'est une dame
    
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
Public Property Let Column(pColumn As Integer)
    prvColumn = pColumn
End Property




'/// PROPRIÉTÉ : Retourne la couleur du pion
'/// PARAMÈTRE : Aucun
'/// RETOUR    : Integer
Public Property Get Row() As Integer
    Row = prvRow
End Property



'/// PROPRIÉTÉ : Modifie la variable privée correspondant à la ligne
'/// PARAMÈTRE : Integer
'/// RETOUR    : Aucun (Let)
Public Property Let Row(pRow As Integer)
    prvRow = pRow
End Property



'/// PROPRIÉTÉ  : Retourne la valeur de la range du pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Range
Public Property Get CurrentRange() As Range

    'Construit une range sur la cellule correspondant a la ligne et a la colonne des variables privées du pion
    Set CurrentRange = Cells(prvRow, prvColumn)
    
End Property



'/// PROPRIÉTÉ  : Retourne la couleur associé a la cellule de la range du pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EColor
Public Property Get Color() As EColor
    Color = prvColor
End Property



'/// PROPRIÉTÉ  : Retourne la couleur opposée a la cellule de la range du pion
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EColor
Public Property Get EnemyColor() As EColor
    
    If Me.Color = EColor.White Or Me.Color = EColor.WhiteMoveStep Then
        EnemyColor = EColor.Black
    ElseIf Me.Color = EColor.Black Or Me.Color = EColor.BlackMoveStep Then
        EnemyColor = EColor.White
    End If
    
End Property



'/// PROPRIÉTÉ  : Change la couleur de la cellule de la range du pion
'/// PARAMÈTRE  : EColor
'/// RETOUR     : Aucun
Public Property Let Color(pColor As EColor)
    
    prvColor = pColor
    
    Select Case pColor
    
        Case EColor.White
            CurrentRange.Font.Color = RGB(255, 255, 255)
            
        Case EColor.Black
            CurrentRange.Font.Color = RGB(0, 0, 0)
            
        Case EColor.BlackMoveStep
            CurrentRange.Font.Color = RGB(150, 50, 50)
            
        Case EColor.WhiteMoveStep
            CurrentRange.Font.Color = RGB(230, 180, 180)
            
    End Select
    
End Property



'/// FONCTION   : Retourne 1 si le pion est noir, -1 s'il est blanc
'///              Cela correspond a la direction vers laquelle le pion a le droit de se déplacer
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Integer
Property Get Orientation() As Integer
    
    If Me.Color = EColor.White Or Me.Color = EColor.WhiteMoveStep Then
        Orientation = -1
    ElseIf Me.Color = EColor.Black Or Me.Color = EColor.BlackMoveStep Then
        Orientation = 1
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne vrai si le pion est une Dame
'/// PARAMÈTRE  : Aucun
'/// retourne   : Boolean
Public Property Get IsQueen() As Boolean
    IsQueen = prvIsQueen
End Property



'/// PROPRIÉTÉ  : Retourne vrai si le pion est une Dame
'/// PARAMÈTRE  : Aucun
'/// retourne   : Boolean
Public Property Let IsQueen(pIsQueen As Boolean)
    
    prvIsQueen = pIsQueen
    
    If pIsQueen Then
        CurrentRange.Value = Chr(169)
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne vrai si le pion est véritablement un pion
'/// PARAMÈTRE  : Aucun
'/// retourne   : Boolean
Public Property Get IsPawn() As Boolean
    IsPawn = prvIsPawn
End Property



'/// PROPRIÉTÉ  : Retourne vrai si le pion est véritablement un pion
'/// PARAMÈTRE  : Aucun
'/// retourne   : Boolean
Public Property Let IsPawn(pIsPawn As Boolean)
    prvIsPawn = pIsPawn
End Property



'/// PROPRIÉTÉ  : Retourne vrai si le pion est une step de déplacement un pion
'/// PARAMÈTRE  : Aucun
'/// retourne   : Boolean
Public Property Get IsStepMove() As Boolean
    If Me.Color = EColor.WhiteMoveStep Or Me.Color = EColor.BlackMoveStep Then
        IsStepMove = True
    End If
End Property



'/// FONCTION   : Transforme le pion en dame s'il est sur la derniere ligne de son orientation
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



'/// PROPRIÉTÉ  : Vérifie si le mouvement est autorisé et, si possible, déplace le pion
'///              Retourne vrai si le pion peut se déplacer
'/// PARAMÈTRE  : Pion (à la position final apres déplacement), Bouléen indiquant si on déplace le pion
'/// RETOUR     : Boolean
Public Function TryMoveTo(pTarget As PawnModel, pDoMove As Boolean) As Boolean
    
    If Me.CanMoveTo(pTarget) Then
        
        TryMoveTo = True
        
        If pDoMove Then
        
            Call MoveTo(pTarget)
        
        End If
        
    End If
    
End Function



'/// PROPRIÉTÉ  : déplace le pion vers la position passé en parametre
'/// PARAMÈTRE  : Pion (à la position final apres déplacement)
'/// RETOUR     : Aucun
Private Sub MoveTo(ByRef pTarget As PawnModel)
    
    Dim move As MoveModel
    Set move = New MoveModel
    Call move.Build(Me, pTarget)
    
    'si c'est une attaque
    If move.IsAttack Then
        'on efface le pion mangé
        move.EnemyKilled.CurrentRange.ClearContents
    End If
    
    'on transmet la valeur (pion ou dame) et la couleur du pion a la cellule de destination
    pTarget.CurrentRange.Value = Me.CurrentRange.Value
    pTarget.Color = Me.Color
    Call pTarget.Build(pTarget.CurrentRange)
    
    'on vérifie si le pion peut devenir une dame
    Call pTarget.TryBecomeQueen
    
    'on efface le pion de sa position initial
    Me.CurrentRange.ClearContents
    
End Sub



'/// PROPRIÉTÉ  : Vérifie si le mouvement est autorisé
'/// PARAMÈTRE  : Pion (à la position final apres déplacement)
'/// RETOUR     : Boolean
Public Function CanMoveTo(pTarget As PawnModel) As Boolean
    
    Dim move As MoveModel
    
    'on initialise le retour de la fonction a false
    CanMoveTo = False
    
    'on instancie un nouvel object déplacement que l'on construit avec le pion actuelle et le pion à la position finale souhaité
    Set move = New MoveModel
    Call move.Build(Me, pTarget)
    
    'si c'est un déplacement et aucun pion allié ne peut attaquer ou que ce déplacement est une attaque
    If (move.IsMove And Not Me.HasAttackingAlly) Or move.IsAttack Then
        CanMoveTo = True
    End If
    
    
End Function



'/// PROPRIÉTÉ  : Retourne true si le pion peut attaquer
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Property Get CanAttack() As Boolean
    
    Dim move As MoveModel
    
    Dim outOfRange As Boolean
    Dim iVector As Integer
    Dim coefVector As Variant
    Dim cardinalDirections As Variant
    Dim Row As Integer, Column As Integer
    Dim Direction As EWindRose
    
    Dim nextCell As Range
    Dim targetCandidate As PawnModel
    Dim board As BoardModel
    
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
        
            Direction = cardinalDirection
            'on récupere les coefficients associés à la direction du déplacement
            coefVector = board.CardinalCoef(Direction)
            
            iVector = 0
            outOfRange = False
            
            'si le pion est une dame
            If IsQueen Then
            
                'Tant qu'on a pas détécter un pion a manger ou la fin du plateau, on boucle
                While Not (outOfRange)
                
                    iVector = iVector + 1
                    
                    Row = Me.Row + iVector * coefVector(0)
                    Column = Me.Column + iVector * coefVector(1)
                    
                    'on récupère les coordonnées de la n-ieme cellule sur la diagonale
                    Set nextCell = Cells(Row, Column)
                    
                    'si cellule sur damier
                    If board.GetSection(nextCell) = ESection.Game Then
                    
                        'on construit l'objet pion
                        Set targetCandidate = New PawnModel
                        Call targetCandidate.Build(nextCell)
                        
                        'si la case est vide et est prete a accueillir un déplacement
                        If Not targetCandidate.IsPawn Then
                            
                            'on construit un objet déplacement a partir du pion actuel et la cellule en cours de test comme destination de mouvement
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
                        
                        'si la case est vide et est prete a accueillir un déplacement
                        If Not targetCandidate.IsPawn Then
                            
                            'on construit un objet déplacement a partir du pion actuel et la cellule en cours de test comme destination de mouvement
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
Property Get CanMove() As Boolean
    
    Dim move As MoveModel
    
    'on initialise le retour de la fonction a false
    CanMove = False
    
    Dim coefVector As Variant
    Dim cardinalDirections As Variant
    Dim Row As Integer, Column As Integer
    Dim Direction As EWindRose
    
    Dim nextCell As Range
    Dim targetCandidate As PawnModel
    Dim board As BoardModel
    
    'si aucun allié ne peut attaquer
    If Me.IsPawn And Not Me.HasAttackingAlly Then
    
        Set targetCandidate = New PawnModel
        Set board = New BoardModel
        Set move = New MoveModel
    
        'on construit un tableau remplit des 4 directions possibles
        cardinalDirections = Array(EWindRose.NorthEast, EWindRose.NorthWest, EWindRose.SouthEast, EWindRose.SouthWest)
        
        'pour chaque direction possible dans le tableau des directions possible
        For Each cardinalDirection In cardinalDirections
        
            Direction = cardinalDirection
            'on récupere les coefficients associés à la direction du déplacement
            coefVector = board.CardinalCoef(Direction)
            
            'si le pion est une dame
            'ou si on est en train de tester une direction au nord pour un pion blanc, sud pour un pion noir
            If Me.IsQueen Or coefVector(0) = Me.Orientation Then
        
                Row = Me.Row + 1 * coefVector(0)
                Column = Me.Column + 1 * coefVector(1)
                
                'on récupère les coordonnées de la n-ieme cellule sur la diagonale
                Set nextCell = Cells(Row, Column)
                
                'si cellule sur damier
                If board.GetSection(nextCell) = ESection.Game Then
                
                    'on construit l'objet pion
                    Set targetCandidate = New PawnModel
                    Call targetCandidate.Build(nextCell)
                    
                    'si la case est vide et est prete a accueillir un déplacement
                    If Not targetCandidate.IsPawn Then
                        
                        'on construit un objet déplacement a partir du pion actuel et la cellule en cours de test comme destination de mouvement
                        Call move.Build(Me, targetCandidate)
                        
                        'si le déplcament est un mouvement
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
Property Get HasAttackingAlly() As Boolean

    Dim pawnList As Variant
    Dim currentPawn As PawnModel
    
    Set currentPawn = New PawnModel
    
    'on initialise le retour de la fonction a false
    HasAttackingAlly = False
    
    'on récupère le tableau des pions alliés
    pawnList = Game.GetPawns(Me.Color)
    
    'pour chaque pion allié dans le tableau de pion allié
    For Each pawn In pawnList
    
        Set currentPawn = pawn
        
        'si le pion allié peut attaquer
        If currentPawn.CanAttack Then
        
            HasAttackingAlly = True
            Exit For
        
        End If
    
    Next pawn


End Property



'/// PROCÉDURE  : Transforme le pion en étape de déplacement sur le plateau
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Sub MarkMoveStep(pawn As PawnModel)

    pawn.CurrentRange.Value = Me.CurrentRange.Value
    
    If Me.Color = EColor.White Or Me.Color = EColor.WhiteMoveStep Then
        pawn.Color = EColor.WhiteMoveStep
    ElseIf Me.Color = EColor.Black Or Me.Color = EColor.BlackMoveStep Then
        pawn.Color = EColor.BlackMoveStep
    End If
        
End Sub
