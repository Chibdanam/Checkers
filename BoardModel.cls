VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/// CONSTRUCTEUR: methode appel� lors de l'instanciation de l'objet
'/// PARAM�TRE   : Aucun (evenement constructeur)
'/// RETOUR      : Aucun (evenement)
Private Sub class_Initialize()
    
End Sub



'/// DESTRUCTEUR: methode appel� lors de la destruction de l'objet
'/// PARAM�TRE  : Aucun (evenement destructeur)
'/// RETOUR     : Aucun (evenemnt)
Private Sub class_Terminate()

End Sub



'/// PROPRI�T�  : Retourne le nom de la section dans laquelle se trouve la Range pass�e en parametre
'/// PARAM�TRE  : Range
'/// RETOUR     : ESection
Public Property Get GetSection(cell As Range) As ESection

    'par defaut, on definie la valeur de retour a Hors limite
    GetSection = ESection.OutOfLimit
    
    'si la cellule est comprise danse une des sections connues du plateau, alors on renvoit la valeur de la section
    If Not (Application.Intersect(cell, Range("Game")) Is Nothing) Then
        GetSection = ESection.Game
    ElseIf Not (Application.Intersect(cell, Range("ConfigPlayer")) Is Nothing) Then
        GetSection = ESection.ConfigPlayer
    ElseIf Not (Application.Intersect(cell, Range("Restart")) Is Nothing) Then
        GetSection = ESection.Restart
    End If

End Property



'/// PROPRI�T�  : Retourne la valeur de l'enumeration EColor associ�e de la valeur �crite sur le damier
'/// PARAM�TRE  : Aucun
'/// RETOUR     : EColor
Public Property Get CurrentTurn() As EColor

    If Range("TurnValue") = "White" Then
        CurrentTurn = EColor.White
    ElseIf Range("TurnValue") = "Black" Then
        CurrentTurn = EColor.Black
    Else
        'Si aucune valeur attendue n'est trouv�, on la parametre a "White" par defaut
        CurrentTurn = EColor.White
        Range("TurnValue") = "White"
    End If
    
End Property



'/// PROC�DURE  : Change la couleur afficher sur le plateau de jeu
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SwitchTurn()

    If Range("TurnValue") <> "You Win" Then
        If CurrentTurn = EColor.Black Then
            CurrentTurn = EColor.White
        ElseIf CurrentTurn = EColor.White Then
            CurrentTurn = EColor.Black
        End If
    End If
    
End Sub



'/// PROPRI�T�  : Retourne
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Property Let CurrentTurn(pColor As EColor)

    If pColor = EColor.Black Then
        Range("TurnValue") = "Black"
    ElseIf pColor = EColor.White Then
        Range("TurnValue") = "White"
    End If
    
End Property



'/// PROPRI�T�  : Retourne la valeur �crite sur le damier puis renvoie la valeur de l'enumeration EConfig associ�e
'/// PARAM�TRE  : Aucun
'/// RETOUR     : EConfig
Public Property Get ConfigPlayer() As EConfig
    
    If Range("ConfigPlayerValue") = "1 Player" Then
        ConfigPlayer = EConfig.SinglePlayer
    ElseIf Range("ConfigPlayerValue").Value = "2 Players" Then
        ConfigPlayer = EConfig.TwoPlayer
    ElseIf Range("ConfigPlayerValue").Value = "Automa" Then
        ConfigPlayer = EConfig.Automate
    Else
        'Si aucune valeur attendue n'est trouv�, on la parametre a "1 Player" par defaut
        ConfigPlayer = SinglePlayer
        Range("ConfigPlayerValue") = "1 Player"
    End If
    
End Property



'/// PROPRI�T�  : Ecris la valeur pass� en parametre sur la plateau de jeu
'/// PARAM�TRE  : EConfig
'/// RETOUR     : Aucun
Public Property Let ConfigPlayer(Config As EConfig)

    If Config = EConfig.SinglePlayer Then
        Range("ConfigPlayer") = "1 Players"
    ElseIf Config = EConfig.TwoPlayer Then
        Range("ConfigPlayer") = "2 Players"
    ElseIf Config = EConfig.Automate Then
        Range("ConfigPlayer") = "Automa"
    End If
        
End Property



'/// PROPRI�T�  : �cris sur le plateau de jeu dans la zone "m�moire" la position du pion pass� en parametre
'/// PARAM�TRE  : PawnModel
'/// RETOUR     : Acun
Public Property Let Memory(pawn As PawnModel)

    Dim Column As String
    Dim Row As String
    
    'permet d'�crire sur le plateau les coordon�es dans le r�f�rentiel du plateau et non le r�f�rentiel d'excel
    Dim offset As Integer
    
    'le damier commencant en "B2", on a un offset de 1 sur la colonne et de 1 sur l ligne par rapport a "A1"
    offset = 1
    
    'afin de r�cuper la lettre correspondant a la colonne, on convertie la valeur de la colonne en charactere ASCII
    Column = Chr(64 + (pawn.CurrentRange.Column - offset))
    'on convertie notre valeur num�ric en chaine de charactere : Integer to String
    Row = CStr(pawn.CurrentRange.Row - offset)
    
    Range("MemoryValue").Value = Column + Row
        
End Property



'/// PROPRI�T�  : Retourne le pion �cris en m�moire sur le plateau de jeu
'/// PARAM�TRE  : Aucun
'/// RETOUR     : PawnModel
Public Property Get Memory() As PawnModel

    Dim offsetPawnPosition As String
    Dim offsetPawnRange As Range
    Dim pawnRange As Range
    Dim offset As Integer
    
    Set Memory = New PawnModel
    
    offset = 1
    
    'si la range "m�moire" n'est pas vie
    If Not IsEmpty(Range("MemoryValue")) Then
        
        'on r�cupere la valeur �crite sur le plateau
        offsetPawnPosition = Range("MemoryValue").Value
        
        'r�cupre la cellule dans le r�f�rentiel du damier
        Set offsetPawnRange = Range(offsetPawnPosition)
        
        'r�cupere la veritable cellule dans le r�f�rentiel d'excel
        Set pawnRange = Cells(offsetPawnRange.Row + offset, offsetPawnRange.Column + offset)
        
        'on construit le pion avec la cellule pr�c�demment r�cup�r�
        Call Memory.Build(pawnRange)
        
    End If
    
End Property



'/// PROPRI�T�  : Ecris sur la plateau de jeu "You Win" si il ne reste qu'un seul joueur
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SearchWinner()

    Dim whiteFound As Boolean, blackFound As Boolean
    Dim pawn As PawnModel
    
    whiteFound = False
    blackFound = False
    
    'pour chaque cellule du damier
    For Each cell In Range("Game")
        
        'on instancie et construie un pion avec la cellule actuelle
        Set pawn = New PawnModel
        Call pawn.Build(Cells(cell.Row, cell.Column))
        
        'si le pion est v�ritablement un pion
        If pawn.IsPawn Then
        
            'si le pion est noir
            If pawn.Color = EColor.Black Then
                blackFound = True
                
            'si le pion est blanc
            Else
                whiteFound = True
            End If
            
            'si on a trouv� un pion blanc et un pion noir
            If whiteFound And blackFound Then
                Exit For
            End If
            
        End If
    Next cell
    
    'si une des couleurs n'est pas trouv�
    If Not whiteFound Or Not blackFound Then
        Range("TurnValue") = "You Win"
    End If
    
End Sub



'/// PROPRI�T�  : Retourne true si "You win" est �cris sur le plateau
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Boolean
Property Get IsWinner() As Boolean

    IsWinner = False
    
    If Range("TurnValue") = "You Win" Then
        IsWinner = True
    End If
    
End Property



'/// PROPRI�T�  : Retourne un tableau comprenant les coefficients associ�s � la direction du d�placement
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Variant (tableau de la forme (x,y))
Property Get CardinalCoef(Direction As EWindRose) As Variant
    
    If Direction = EWindRose.SouthEast Then
        CardinalCoef = Array(1, 1)
    ElseIf Direction = EWindRose.NorthEast Then
        CardinalCoef = Array(-1, 1)
    ElseIf Direction = EWindRose.SouthWest Then
        CardinalCoef = Array(1, -1)
    ElseIf Direction = EWindRose.NorthWest Then
        CardinalCoef = Array(-1, -1)
    End If
    
End Property



