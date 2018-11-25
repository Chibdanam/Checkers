VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PawnModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/// VARIABLE PRIV�E : Ligne
Private prvRow As Integer
'/// VARIABLE PRIV�E : Colonne
Private prvColumn As Integer
'/// VARIABLE PRIV�E : Couleur
Private prvColor As EColor
'/// VARIABLE PRIV�E : Couleur
Private prvIsPawn As Boolean
'/// VARIABLE PRIV�E : Colonne
Private prvIsQueen As Boolean



'/// CONSTRUCTEUR: methode appel� lors de l'instanciation de l'objet
'/// PARAM�TRE   : Aucun (evenement constructeur)
'/// RETOUR      : Aucun (evenement)
Private Sub class_Initialize()
    'on initialise les variables avec les valeurs de la cellule A1
    prvRow = 1
    prvColumn = 1
End Sub



'/// DESTRUCTEUR: methode appel� lors de la destruction de l'objet
'/// PARAM�TRE  : Aucun (evenement destructeur)
'/// RETOUR     : Aucun (evenemnt)
Private Sub class_Terminate()
    
End Sub



'/// PROC�DURE  : parametre tous les attributs du pion � partir d'une range
'/// PARAM�TRE  : Range
'/// RETOUR     : Aucun
Public Sub Build(ByVal cell As Range)
    
    prvRow = cell.Row
    prvColumn = cell.Column

        '/// On d�finit la variable priv� de couleur du pion
    
    If cell.Font.Color = RGB(255, 255, 255) Then
        prvColor = EColor.White
    ElseIf cell.Font.Color = RGB(0, 0, 0) Then
        prvColor = EColor.Black
    End If
    
    '/// On d�finit si la cellule contient un pion
    
    If Not IsEmpty(cell) And (Me.Color = EColor.Black Or Me.Color = EColor.White) Then
        prvIsPawn = True
    Else
        prvIsPawn = False
    End If
    
    '/// On d�finit la variable priv� d'�tat Dame du pion, si la valeur de la cellule est � alors c'est une dame
    
    If InStr(Chr(169), cell.Value) Then
        prvIsQueen = True
    Else
        prvIsQueen = False
    End If
    
End Sub



'/// PROPRI�T� : Retourne la colonne du pion
'/// PARAM�TRE : Aucun
'/// RETOUR    : Integer
Public Property Get Column() As Integer
    Column = prvColumn
End Property



'/// PROPRI�T� : Modifie la variable priv�e correspondant � la colonne
'/// PARAM�TRE : Integer
'/// RETOUR    : Aucun (Let)
Public Property Let Column(pColumn As Integer)
    prvColumn = pColumn
End Property




'/// PROPRI�T� : Retourne la couleur du pion
'/// PARAM�TRE : Aucun
'/// RETOUR    : Integer
Public Property Get Row() As Integer
    Row = prvRow
End Property



'/// PROPRI�T� : Modifie la variable priv�e correspondant � la ligne
'/// PARAM�TRE : Integer
'/// RETOUR    : Aucun (Let)
Public Property Let Row(pRow As Integer)
    prvRow = pRow
End Property



'/// PROPRI�T�  : Retourne la valeur de la range du pion
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Range
Public Property Get CurrentRange() As Range

    'Construit une range sur la cellule correspondant a la ligne et a la colonne des variables priv�es du pion
    Set CurrentRange = Cells(prvRow, prvColumn)
    
End Property



'/// PROPRI�T�  : Retourne la couleur associ� a la cellule de la range du pion
'/// PARAM�TRE  : Aucun
'/// RETOUR     : EColor
Public Property Get Color() As EColor
   Color = prvColor
End Property



'/// PROPRI�T�  : Change la couleur de la cellule de la range du pion
'/// PARAM�TRE  : EColor
'/// RETOUR     : Aucun
Public Property Let Color(pColor As EColor)
    
    prvColor = pColor
    
    Select Case pColor
    
        Case EColor.White
            CurrentRange.Font.Color = RGB(255, 255, 255)
            
        Case EColor.Black
            CurrentRange.Font.Color = RGB(0, 0, 0)
            
    End Select
    
End Property



'/// PROPRI�T�  : Retourne la couleur oppos�e a la cellule de la range du pion
'/// PARAM�TRE  : Aucun
'/// RETOUR     : EColor
Public Property Get EnemyColor() As EColor
    
    If Me.Color = EColor.White Then
        EnemyColor = EColor.Black
    ElseIf Me.Color = EColor.Black Then
        EnemyColor = EColor.White
    End If
    
End Property



'/// FONCTION   : Retourne 1 si le pion est noir, -1 s'il est blanc
'///              Cela correspond a la direction vers laquelle le pion a le droit de se d�placer
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Integer
Property Get Orientation() As Integer
    
    If Me.Color = EColor.White Then
        Orientation = -1
    ElseIf Me.Color = EColor.Black Then
        Orientation = 1
    End If
    
End Property



'/// PROPRI�T�  : Retourne vrai si le pion est une Dame
'/// PARAM�TRE  : Aucun
'/// retourne   : Boolean
Public Property Get IsQueen() As Boolean
    IsQueen = prvIsQueen
End Property



'/// PROPRI�T�  : Retourne vrai si le pion est une Dame
'/// PARAM�TRE  : Aucun
'/// retourne   : Boolean
Public Property Let IsQueen(pIsQueen As Boolean)
    
    prvIsQueen = pIsQueen
    
    If pIsQueen Then
        CurrentRange.Value = Chr(169)
    End If
    
End Property



'/// PROPRI�T�  : Retourne vrai si le pion est v�ritablement un pion
'/// PARAM�TRE  : Aucun
'/// retourne   : Boolean
Public Property Get IsPawn() As Boolean
    IsPawn = prvIsPawn
End Property



'/// PROPRI�T�  : Retourne vrai si le pion est v�ritablement un pion
'/// PARAM�TRE  : Aucun
'/// retourne   : Boolean
Public Property Let IsPawn(pIsPawn As Boolean)
    prvIsPawn = pIsPawn
End Property



'/// FONCTION   : Transforme le pion en dame s'il est sur la derniere ligne de son orientation
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub TryBecomeQueen()

    If Not IsQueen Then
    
        If Me.Color = EColor.Black And Me.Row = 9 Or _
           Me.Color = EColor.White And Me.Row = 2 Then

            Me.IsQueen = True
                
        End If
    End If
    
End Sub



'/// PROPRI�T�  : V�rifie si le mouvement est autoris� et, si possible, d�place le pion
'///              Retourne vrai si le pion peut se d�placer
'/// PARAM�TRE  : Pion (� la position final apres d�placement), Boul�en indiquant si on d�place le pion
'/// RETOUR     : Boolean
Public Function TryMoveTo(pTarget As PawnModel, pDoMove As Boolean) As Boolean
    
    If Me.CanMoveTo(pTarget) Then
        
        TryMoveTo = True
        
        If pDoMove Then
        
            Call MoveTo(pTarget)
        
        End If
        
    End If
    
End Function



'/// PROPRI�T�  : d�place le pion vers la position pass� en parametre
'/// PARAM�TRE  : Pion (� la position final apres d�placement)
'/// RETOUR     : Aucun
Private Sub MoveTo(ByRef pTarget As PawnModel)
    
    Dim move As MoveModel
    Set move = New MoveModel
    Call move.Build(Me, pTarget)
    
    'si c'est une attaque
    If move.IsAttack Then
        'on efface le pion mang�
        move.EnemyKilled.CurrentRange.ClearContents
    End If
    
    'on transmet la valeur (pion ou dame) et la couleur du pion a la cellule de destination
    pTarget.CurrentRange.Value = Me.CurrentRange.Value
    pTarget.Color = Me.Color
    Call pTarget.Build(pTarget.CurrentRange)
    
    'on v�rifie si le pion peut devenir une dame
    Call pTarget.TryBecomeQueen
    
    'on efface le pion de sa position initial
    Me.CurrentRange.ClearContents
    
End Sub



'/// PROPRI�T�  : V�rifie si le mouvement est autoris�
'/// PARAM�TRE  : Pion (� la position final apres d�placement)
'/// RETOUR     : Boolean
Public Property Get CanMoveTo(pTarget As PawnModel) As Boolean
    
    Dim move As MoveModel
    
    'on initialise le retour de la fonction a false
    CanMoveTo = False
    
    'on instancie un nouvel object d�placement que l'on construit avec le pion actuelle et le pion � la position finale souhait�
    Set move = New MoveModel
    Call move.Build(Me, pTarget)
    
    'si c'est un d�placement et aucun pion alli� ne peut attaquer ou que ce d�placement est une attaque
    If (move.IsMove And Not Me.HasAttackingAlly) Or move.IsAttack Then
        CanMoveTo = True
    End If
    
    
End Property



'/// PROPRI�T�  : Retourne true si le pion peut attaquer
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get CanAttack() As Boolean
    
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

    'si le pion est v�ritablement un pion
    If Me.IsPawn Then
    
        'on construit un tableau remplit des 4 directions possibles
        cardinalDirections = Array(EWindRose.NorthEast, EWindRose.NorthWest, EWindRose.SouthEast, EWindRose.SouthWest)
        
        'pour chaque direction possible dans le tableau des directions possible
        For Each cardinalDirection In cardinalDirections
        
            Direction = cardinalDirection
            'on r�cupere les coefficients associ�s � la direction du d�placement
            coefVector = board.CardinalCoef(Direction)
            
            iVector = 0
            outOfRange = False
            
            'si le pion est une dame
            If IsQueen Then
            
                'Tant qu'on a pas d�t�cter un pion a manger ou la fin du plateau, on boucle
                While Not (outOfRange)
                
                    iVector = iVector + 1
                    
                    Row = Me.Row + iVector * coefVector(0)
                    Column = Me.Column + iVector * coefVector(1)
                    
                    'on r�cup�re les coordonn�es de la n-ieme cellule sur la diagonale
                    Set nextCell = Cells(Row, Column)
                    
                    'si cellule sur damier
                    If board.GetSection(nextCell) = ESection.Game Then
                    
                        'on construit l'objet pion
                        Set targetCandidate = New PawnModel
                        Call targetCandidate.Build(nextCell)
                        
                        'si la case est vide et est prete a accueillir un d�placement
                        If Not targetCandidate.IsPawn Then
                            
                            'on construit un objet d�placement a partir du pion actuel et la cellule en cours de test comme destination de mouvement
                            Call move.Build(Me, targetCandidate)
                            
                            'si le d�placement est une attaque
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
                        
                        'si la case est vide et est prete a accueillir un d�placement
                        If Not targetCandidate.IsPawn Then
                            
                            'on construit un objet d�placement a partir du pion actuel et la cellule en cours de test comme destination de mouvement
                            Call move.Build(Me, targetCandidate)
                            
                            'si le d�placement est une attaque
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



'/// PROPRI�T�  : Retourne true si le pion peut se d�placer
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get CanMove() As Boolean
    
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
    
    'si aucun alli� ne peut attaquer
    If Me.IsPawn And Not Me.HasAttackingAlly Then
    
        Set targetCandidate = New PawnModel
        Set board = New BoardModel
        Set move = New MoveModel
    
        'on construit un tableau remplit des 4 directions possibles
        cardinalDirections = Array(EWindRose.NorthEast, EWindRose.NorthWest, EWindRose.SouthEast, EWindRose.SouthWest)
        
        'pour chaque direction possible dans le tableau des directions possible
        For Each cardinalDirection In cardinalDirections
        
            Direction = cardinalDirection
            'on r�cupere les coefficients associ�s � la direction du d�placement
            coefVector = board.CardinalCoef(Direction)
            
            'si le pion est une dame
            'ou si on est en train de tester une direction au nord pour un pion blanc, sud pour un pion noir
            If Me.IsQueen Or coefVector(0) = Me.Orientation Then
        
                Row = Me.Row + 1 * coefVector(0)
                Column = Me.Column + 1 * coefVector(1)
                
                'on r�cup�re les coordonn�es de la n-ieme cellule sur la diagonale
                Set nextCell = Cells(Row, Column)
                
                'si cellule sur damier
                If board.GetSection(nextCell) = ESection.Game Then
                
                    'on construit l'objet pion
                    Set targetCandidate = New PawnModel
                    Call targetCandidate.Build(nextCell)
                    
                    'si la case est vide et est prete a accueillir un d�placement
                    If Not targetCandidate.IsPawn Then
                        
                        'on construit un objet d�placement a partir du pion actuel et la cellule en cours de test comme destination de mouvement
                        Call move.Build(Me, targetCandidate)
                        
                        'si le d�plcament est un mouvement
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



'/// PROPRI�T�  : Retourne true si un alli� peut attaquer
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get HasAttackingAlly() As Boolean

    Dim pawnList As Variant
    Dim currentPawn As PawnModel
    
    Set currentPawn = New PawnModel
    
    'on initialise le retour de la fonction a false
    HasAttackingAlly = False
    
    'on r�cup�re le tableau des pions alli�s
    pawnList = Tools.GetPawns(Me.Color)
    
    'pour chaque pion alli� dans le tableau de pion alli�
    For Each pawn In pawnList
    
        Set currentPawn = pawn
        
        'si le pion alli� peut attaquer
        If currentPawn.CanAttack Then
        
            HasAttackingAlly = True
            Exit For
        
        End If
    
    Next pawn


End Property
