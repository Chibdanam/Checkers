VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'/// CONSTRUCTEUR: méthode appelé lors de l'instanciation de l'objet
'/// PARAMÈTRE   : Aucun
'/// RETOUR      : Aucun
Private Sub class_Initialize()
End Sub



'/// DESTRUCTEUR: méthode appelé lors de la destruction de l'objet
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Private Sub class_Terminate()
End Sub



'/// PROPRIÉTÉ  : Retourne le nom de la section dans laquelle se trouve la Range passée en paramètre
'/// PARAMÈTRE  : Range
'/// RETOUR     : ESection
Public Property Get GetSection(ByVal pCell As Range) As ESection

    'par défaut, on définit la valeur de retour à Hors limite
    GetSection = ESection.OutOfLimit
    
    'si la cellule est comprise danse une des sections connues du plateau, alors on renvoi la valeur de la section
    If Not (Application.Intersect(pCell, Range("Game")) Is Nothing) Then
        GetSection = ESection.Game
    ElseIf Not (Application.Intersect(pCell, Range("ConfigPlayer")) Is Nothing) Then
        GetSection = ESection.ConfigPlayer
    ElseIf Not (Application.Intersect(pCell, Range("Restart")) Is Nothing) Then
        GetSection = ESection.Restart
    End If

End Property



'/// PROPRIÉTÉ  : Retourne la valeur de l'énumération EColor associée de la valeur écrite sur le damier
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EColor
Public Property Get CurrentTurn() As EColor

    If Range("TurnValue") = EnumString(EColor.White) Then
        CurrentTurn = EColor.White
    ElseIf Range("TurnValue") = EnumString(EColor.Black) Then
        CurrentTurn = EColor.Black
    Else
        'Si aucune valeur attendue n'est trouvée, on la paramètre à "White" par défaut
        CurrentTurn = EColor.White
        Range("TurnValue") = EnumString(EColor.White)
    End If
    
End Property



'/// PROCÉDURE  : Change la couleur affichée sur le plateau de jeu
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SwitchTurn()

    If Range("TurnValue") <> "You Win" Then
        If Me.CurrentTurn = EColor.Black Then
            Me.CurrentTurn = EColor.White
        ElseIf Me.CurrentTurn = EColor.White Then
            Me.CurrentTurn = EColor.Black
        End If
    End If
    
End Sub



'/// PROPRIÉTÉ  : Retourne la couleur du joueur du tour en cours
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Property Let CurrentTurn(ByVal pColor As EColor)
    Range("TurnValue") = EnumString(pColor)
End Property



'/// PROPRIÉTÉ  : Retourne la configuration actuelle du plateau jeu (1 joueur, 2 joueurs, ia vs ia)
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EConfig
Public Property Get ConfigPlayer() As EConfig
    
    If Range("ConfigPlayerValue") = EnumString(EConfig.SinglePlayer) Then
        ConfigPlayer = EConfig.SinglePlayer
    ElseIf Range("ConfigPlayerValue").Value = EnumString(EConfig.TwoPlayers) Then
        ConfigPlayer = EConfig.TwoPlayers
    ElseIf Range("ConfigPlayerValue").Value = EnumString(EConfig.Automate) Then
        ConfigPlayer = EConfig.Automate
    Else
        'Si aucune valeur attendue n'est trouvée, on la paramètre a "1 Player" par défaut
        ConfigPlayer = EConfig.SinglePlayer
        Range("ConfigPlayerValue") = EnumString(EConfig.SinglePlayer)
    End If
    
End Property



'/// PROPRIÉTÉ  : Modifie la configuration du plateau de jeu
'/// PARAMÈTRE  : EConfig
'/// RETOUR     : Aucun
Public Property Let ConfigPlayer(ByVal pConfig As EConfig)
        Range("ConfigPlayer") =  EnumString(pConfig)
End Property



'/// PROPRIÉTÉ  : écris sur le plateau de jeu dans la zone "mémoire" la position du pion passé en paramètre
'/// PARAMÈTRE  : PawnModel
'/// RETOUR     : Aucun
Public Property Let Memory(ByVal pPawn As PawnModel)
Dim Column As String
Dim Row As String
Dim offset As Integer

    'permet d'écrire sur le plateau les coordonnées dans le référentiel du plateau et non le référentiel d’Excel
    'le damier commençant en "B2", on a un offset de 1 sur la colonne et de 1 sur la ligne par rapport à "A1"
    offset = 1
    
    'afin de récupérer la lettre correspondant à la colonne, on convertie la valeur de la colonne en caractère ASCII
    Column = Chr(64 + (pPawn.CurrentRange.Column - offset))
    'on convertie notre valeur numérique en chaine de caractère : Integer to String
    Row = CStr(pPawn.CurrentRange.Row - offset)
    
    Range("MemoryValue").Value = Column + Row
        
End Property



'/// PROPRIÉTÉ  : Retourne le pion écris en mémoire sur le plateau de jeu
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : PawnModel
Public Property Get Memory() As PawnModel
Dim offsetPawnPosition As String
Dim offsetPawnRange As Range
Dim pawnRange As Range
Dim offset As Integer
    
    Set Memory = New PawnModel
    
    offset = 1
    
    'si la range "mémoire" n'est pas vie
    If Not IsEmpty(Range("MemoryValue")) Then
        
        'on récupère la valeur écrite sur le plateau
        offsetPawnPosition = Range("MemoryValue").Value
        
        'récupère la cellule dans le référentiel du damier
        Set offsetPawnRange = Range(offsetPawnPosition)
        
        'récupère la véritable cellule dans le référentiel d’Excel
        Set pawnRange = Cells(offsetPawnRange.Row + offset, offsetPawnRange.Column + offset)
        
        'on construit le pion avec la cellule précédemment récupéré
        Call Memory.Build(pawnRange)
        
    End If
    
End Property



'/// PROCÉDURE  : Ecris sur le plateau de jeu "You Win" s’il ne reste qu'un seul joueur
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SearchWinner()
Dim whiteFound As Boolean
Dim blackFound As Boolean
Dim pawn As PawnModel
Dim cell As Range
    
    whiteFound = False
    blackFound = False
    
    'pour chaque cellule du damier
    For Each cell In Range("Game")
        
        'on instancie et construit un pion avec la cellule actuelle
        Set pawn = New PawnModel
        Call pawn.Build(cell)
        
        'si le pion est véritablement un pion
        If pawn.IsPawn Then
        
            'si le pion est noir
            If pawn.Color = EColor.Black Then
                blackFound = True
                
            'si le pion est blanc
            Else
                whiteFound = True
            End If
            
            'si on a trouvé un pion blanc et un pion noir
            If whiteFound And blackFound Then
                Exit For
            End If
            
        End If
    Next cell
    
    'si une des couleurs n'est pas trouvé
    If Not whiteFound Or Not blackFound Then
        Range("TurnValue") = "You Win"
    End If
    
End Sub



'/// PROPRIÉTÉ  : Retourne true si "You win" est écris sur le plateau
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Property Get IsWinner() As Boolean

    IsWinner = False
    
    If Range("TurnValue") = "You Win" Then
        IsWinner = True
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne un tableau comprenant les coefficients associés à la direction du déplacement
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Variant (tableau de la forme (x,y))
Property Get CardinalCoef(ByVal pDirection As EWindRose) As Variant
    
    If pDirection = EWindRose.SouthEast Then
        CardinalCoef = Array(1, 1)
    ElseIf pDirection = EWindRose.NorthEast Then
        CardinalCoef = Array(-1, 1)
    ElseIf pDirection = EWindRose.SouthWest Then
        CardinalCoef = Array(1, -1)
    ElseIf pDirection = EWindRose.NorthWest Then
        CardinalCoef = Array(-1, -1)
    End If
    
End Property
